Option Explicit
Dim shell, fs, folder, executable, command
Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
folder = fs.GetParentFolderName(WScript.ScriptFullName)
executable = fs.BuildPath(folder, "runtime\pythonw.exe")
If Not fs.FileExists(executable) Then
  MsgBox "Use Start Recorder.bat to set up the source version, or download the Windows package.", 48, "File Change Log"
  WScript.Quit 1
End If
command = Chr(34) & executable & Chr(34) & " " & Chr(34) & fs.BuildPath(folder, "run.py") & Chr(34) & " --background"
shell.Run command, 0, False
