"""Bounded, read-only extraction of workbook pivot/filter definitions.

Saved OOXML is evidence of a *state*, never evidence of an action. Unsupported
structures are flagged explicitly and never interpreted as complete/replayable.
"""
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
MAX_PART = 8 * 1024 * 1024
MAX_PARTS = 20000
MAX_PIVOTS = 200
MAX_FIELDS = 1000


def _part(name):
    name = name.replace("\\", "/")
    if not name or name.startswith("/") or "\\" in name:
        raise ValueError("Unusable OOXML part path")
    normalized = posixpath.normpath(name)
    if normalized == ".." or normalized.startswith("../"):
        raise ValueError("OOXML relationship escapes workbook")
    return normalized


def _target(origin, raw):
    if not raw:
        raise ValueError("OOXML relationship missing target")
    if raw.startswith("/"):
        return _part(raw.lstrip("/"))
    return _part(posixpath.join(posixpath.dirname(origin), raw))


def _xml(z, part):
    part = _part(part)
    info = z.getinfo(part)
    if info.file_size > MAX_PART:
        raise ValueError("OOXML part exceeds parser size limit")
    value = z.read(part)
    if b"<!DOCTYPE" in value.upper() or b"<!ENTITY" in value.upper():
        raise ValueError("DTD/entity declarations are prohibited")
    return ET.fromstring(value)


def _relationships(z, origin):
    location = posixpath.join(posixpath.dirname(origin), "_rels",
                              posixpath.basename(origin) + ".rels")
    if location not in z.namelist():
        return {}
    out = {}
    for element in _xml(z, location):
        if element.get("TargetMode", "").lower() == "external":
            continue
        rid = element.get("Id")
        if not rid:
            continue
        out[rid] = (element.get("Type", "").rsplit("/", 1)[-1],
                    _target(origin, element.get("Target")))
    return out


def _elements(node, name):
    child = node.find(MAIN + name) if node is not None else None
    return list(child) if child is not None else []


def _field_name(fields, raw):
    try:
        ix = int(raw)
        if 0 <= ix < len(fields):
            return fields[ix]
    except (TypeError, ValueError):
        pass
    return None


def _axis(node, fields, name, warnings):
    result = []
    for item in _elements(node, name):
        index = item.get("x")
        if index == "-2":  # Data-field axis marker, not a cache field.
            result.append("[Values]")
            continue
        field = _field_name(fields, index)
        if field is None:
            warnings.append("unresolved_" + name + "_field")
        else:
            result.append(field)
    return result


def extract_workbook_logic(path):
    """Read OOXML only, with per-pivot error isolation and explicit completeness.

    Returns a serializable bounded object suitable for a capture event. Parsing
    does not open Excel, calculate formulas, refresh external data, or edit files.
    """
    output = dict(format="ooxml", pivots=[], worksheet_filters=[], table_filters=[],
                  warnings=[], completeness="complete_for_supported_structures",
                  reconstruction_verified=False)
    if Path(path).suffix.lower() not in (".xlsx", ".xlsm"):
        return dict(output, completeness="unsupported_workbook_format",
                    warnings=["Only saved .xlsx/.xlsm supported"])
    with zipfile.ZipFile(path) as z:
        if len(z.infolist()) > MAX_PARTS:
            raise ValueError("OOXML archive contains too many parts")
        files = set(z.namelist())
        if "xl/workbook.xml" not in files:
            raise ValueError("Missing workbook.xml")
        workbook = _xml(z, "xl/workbook.xml")
        rels = _relationships(z, "xl/workbook.xml")
        sheets = workbook.find(MAIN + "sheets")
        if sheets is None:
            raise ValueError("Workbook has no sheet list")
        if len(sheets) > 1000:
            raise ValueError("Workbook has too many sheets for logic extraction")
        for sheet in sheets:
            name = sheet.get("name", "")
            rel = rels.get(sheet.get(REL + "id"))
            if not rel or rel[0] != "worksheet" or rel[1] not in files:
                output["warnings"].append("unsupported_or_missing_sheet:" + name[:80])
                continue
            part = rel[1]
            ws = _xml(z, part)
            auto = ws.find(MAIN + "autoFilter")
            sort = ws.find(MAIN + "sortState")
            hidden = sum(x.get("hidden") in ("1", "true") for x in ws.iter(MAIN + "row"))
            if auto is not None or sort is not None or hidden:
                output["worksheet_filters"].append(
                    dict(sheet=name, reference=auto.get("ref") if auto is not None else None,
                         criteria=_criteria(auto), hidden_rows=hidden,
                         has_sort_state=(sort is not None or
                                         (auto is not None and auto.find(MAIN+"sortState") is not None))))
            ws_rels = _relationships(z, part)
            for kind, pivot_part in ws_rels.values():
                if kind != "pivotTable":
                    continue
                if len(output["pivots"]) >= MAX_PIVOTS:
                    raise ValueError("Pivot definition count exceeds limit")
                try:
                    pivot = _pivot(z, pivot_part, name, files)
                    output["pivots"].append(pivot)
                    if pivot["completeness"] != "supported_subset":
                        output["warnings"].append("partial_pivot:" + name + "/" + pivot["name"][:80])
                except (KeyError, IndexError, ValueError, ET.ParseError) as exc:
                    output["warnings"].append("pivot_read_failed:" + name[:80] +
                                               ":" + type(exc).__name__)
            for kind, table_part in ws_rels.values():
                if kind != "table":
                    continue
                try:
                    table = _xml(z, table_part)
                    filtered = table.find(MAIN + "autoFilter")
                    if filtered is not None:
                        output["table_filters"].append(
                            dict(sheet=name, table=table.get("name"),
                                 reference=filtered.get("ref"), criteria=_criteria(filtered),
                                 has_sort_state=filtered.find(MAIN+"sortState") is not None))
                except (KeyError, ValueError, ET.ParseError):
                    output["warnings"].append("table_filter_read_failed:" + name[:80])
        names = [p.lower() for p in files]
        if any(p.startswith("xl/slicers/") or p.startswith("xl/slicercaches/") for p in names):
            output["warnings"].append("slicers_not_fully_reconstructible")
        if any(p.startswith("xl/model/") for p in names):
            output["warnings"].append("data_model_not_supported")
        if any(p.startswith("xl/connections") or "microsoft.mashup" in p for p in names):
            output["warnings"].append("external_connections_not_replayed")
    if output["warnings"]:
        output["completeness"] = "partial_or_unsupported"
    return output


def _criteria(auto):
    if auto is None:
        return []
    result = []
    for col in auto.findall(MAIN + "filterColumn")[:500]:
        filters = [f.get("val") for f in col.iter(MAIN + "filter")][:500]
        customs = [dict(operator=f.get("operator", "equal"), value=f.get("val"))
                   for f in col.iter(MAIN + "customFilter")][:500]
        other = [el.tag.rsplit("}", 1)[-1] for el in col
                 if el.tag not in (MAIN+"filters", MAIN+"customFilters")]
        result.append(dict(column_index=col.get("colId"), values=filters,
                           custom=customs, unsupported_criteria=other[:20],
                           completeness="partial" if other else "supported_subset"))
    return result


def _pivot(z, part, sheet, files):
    root = _xml(z, part)
    rels = _relationships(z, part)
    definitions = [p for k, p in rels.values() if k == "pivotCacheDefinition"]
    warnings = []
    names = []
    source = {"type": "unknown"}
    if definitions and definitions[0] in files:
        cache = _xml(z, definitions[0])
        fields = _elements(cache, "cacheFields")
        if len(fields) > MAX_FIELDS:
            raise ValueError("Too many pivot cache fields")
        names = [x.get("name", "")[:160] for x in fields]
        cache_source = cache.find(MAIN+"cacheSource")
        if cache_source is not None:
            source = {"type": cache_source.get("type", "unknown")}
            worksheet = cache_source.find(MAIN+"worksheetSource")
            if worksheet is not None:
                source.update(sheet=worksheet.get("sheet"),
                              reference=worksheet.get("ref"),
                              table_or_name=worksheet.get("name"))
            if source["type"] != "worksheet":
                warnings.append("non_worksheet_pivot_source")
    else:
        warnings.append("missing_pivot_cache_definition")
    location = root.find(MAIN+"location")
    rows = _axis(root, names, "rowFields", warnings)
    columns = _axis(root, names, "colFields", warnings)
    page_filters = []
    for item in _elements(root, "pageFields"):
        field = _field_name(names, item.get("fld"))
        if field is None:
            warnings.append("unresolved_page_filter")
        page_filters.append(dict(field=field, selected_item_index=item.get("item")))
    values = []
    for item in _elements(root, "dataFields"):
        field = _field_name(names, item.get("fld"))
        if field is None and item.get("fld") != "-2":
            warnings.append("unresolved_data_field")
        values.append(dict(field=field, aggregation=item.get("subtotal", "sum"),
                           label=item.get("name")))
    hidden_counts = {}
    for ix, field in enumerate(_elements(root, "pivotFields")):
        if ix >= len(names):
            warnings.append("pivot_field_count_mismatch")
            break
        hidden = sum(item.get("h") in ("1", "true") for item in field.iter(MAIN+"item"))
        if hidden:
            hidden_counts[names[ix]] = hidden
    filter_count = len(_elements(root, "filters"))
    if filter_count:
        warnings.append("pivot_advanced_filters_not_reconstructed")
    return dict(sheet=sheet, name=(root.get("name") or "")[:160],
                location=location.get("ref") if location is not None else None,
                source=source, rows=rows, columns=columns,
                page_filters=page_filters, values=values,
                fields_with_hidden_items=hidden_counts,
                advanced_filter_count=filter_count,
                completeness="partial" if warnings else "supported_subset",
                warnings=warnings)


def compare_workbook_logic(before, after):
    """Compare saved states. Never label unobserved actions as user operations."""
    def keyed(data):
        return {(p["sheet"], p["name"]): p for p in data.get("pivots", [])}
    a, b = keyed(before), keyed(after)
    changes = []
    for key in sorted(set(a) | set(b)):
        if key not in a:
            state = "first_observed"
        elif key not in b:
            state = "no_longer_present"
        elif a[key] != b[key]:
            state = "definition_changed"
        else:
            continue
        changes.append(dict(sheet=key[0], name=key[1], observation=state,
                            before=a.get(key), after=b.get(key)))
    return dict(pivot_changes=changes[:100], pivot_changes_omitted=max(0, len(changes)-100),
                before_completeness=before.get("completeness"),
                after_completeness=after.get("completeness"))
