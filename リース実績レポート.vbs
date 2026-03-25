' コマンドウィンドウを表示せずにリース実績レポートを起動
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
ScriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
AppDir = ScriptDir & "\lease-report-app"

WshShell.CurrentDirectory = AppDir
' 第2引数 0 = ウィンドウ非表示, 第3引数 False = 完了を待たない
WshShell.Run "cmd /c npx electron .", 0, False
