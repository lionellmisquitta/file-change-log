#!/usr/bin/env python3
"""Generate genuine XLSX pivot-operation outputs through LibreOffice UNO.

Runs using distro /usr/bin/python3 with python3-uno, connected to a fresh
headless soffice socket. A real spreadsheet engine performs the operations;
no synthetic pivot XML is authored by this driver. This does NOT invoke Excel
COM or establish that Excel raised PivotTableUpdate.
"""
import json
import os
import sys
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, PAGE, DATA
from com.sun.star.sheet.GeneralFunction import SUM, COUNT


def prop(name, value):
    result = PropertyValue()
    result.Name = name
    result.Value = value
    return result


def main(input_file, output_dir, port):
    ctx = uno.getComponentContext()
    resolver = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", ctx)
    remote = resolver.resolve(
        "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % int(port))
    desktop = remote.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",remote)
    document = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(os.path.abspath(input_file)),
        "_blank", 0, (prop("Hidden", True), prop("ReadOnly", False)))
    if document is None:
        raise RuntimeError("LibreOffice could not open source XLSX")
    target = Path(output_dir).resolve()
    target.mkdir(parents=True, exist_ok=True)
    emitted = []
    try:
        source = document.Sheets.getByName("SourceData")
        last_row = source.getCellRangeByName("A1:D1").getDataArray()[0] # Header verification.
        assert last_row[:4] == ("Region", "Product", "Month", "Revenue"), last_row
        document.Sheets.insertNewByName("PivotTest", document.Sheets.Count)
        sheet = document.Sheets.getByName("PivotTest")
        pivots = sheet.DataPilotTables
        descriptor = pivots.createDataPilotDescriptor()
        descriptor.SourceRange = source.getCellRangeByName("A1:D49").RangeAddress
        fields = descriptor.DataPilotFields
        fields.getByName("Region").Orientation = ROW
        measure = fields.getByName("Revenue")
        measure.Orientation = DATA
        measure.Function = SUM
        pivots.insertNewByName("RevenuePivot", sheet.getCellByPosition(0, 0).CellAddress, descriptor)

        def save(number, label):
            dest = target / ("%02d_%s.xlsx" % (number, label))
            document.storeAsURL(uno.systemPathToFileUrl(str(dest)), (
                prop("FilterName", "Calc MS Excel 2007 XML"),
                prop("Overwrite", True)))
            if not dest.exists() or dest.stat().st_size < 1000:
                raise RuntimeError("Pivot stage not saved: %s" % dest)
            emitted.append(str(dest))
        save(1, "pivot_created_sum")
        # LibreOffice does not reliably recalculate a pivot when Function is
        # mutated on a previously inserted DataPilotFields object. Rebuild the
        # native DataPilot descriptor for each saved configuration instead of
        # claiming a Count operation when the output still sums Revenue.
        def configure(aggregate, with_product=False, with_month=False):
            modified = pivots.createDataPilotDescriptor()
            modified.SourceRange = source.getCellRangeByName("A1:D49").RangeAddress
            modified_fields = modified.DataPilotFields
            modified_fields.getByName("Region").Orientation = ROW
            metric = modified_fields.getByName("Revenue")
            metric.Orientation = DATA
            metric.Function = aggregate
            if with_product:
                modified_fields.getByName("Product").Orientation = COLUMN
            if with_month:
                modified_fields.getByName("Month").Orientation = PAGE
            pivots.removeByName("RevenuePivot")
            pivots.insertNewByName(
                "RevenuePivot", sheet.getCellByPosition(0, 0).CellAddress, modified)

        configure(COUNT)
        save(2, "aggregation_changed_count")
        configure(COUNT, with_product=True)
        save(3, "product_column_added")
        configure(COUNT, with_product=True, with_month=True)
        save(4, "month_page_filter_added")
        pivots.removeByName("RevenuePivot")
        save(5, "pivot_deleted")
        print(json.dumps({"engine": "libreoffice_uno", "operations": emitted},
                         sort_keys=True), flush=True)
    finally:
        document.close(True)


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit("usage: uno_pivot_operations.py SOURCE.xlsx OUTPUT_DIR PORT")
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]))
