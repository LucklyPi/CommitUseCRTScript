
' Version		1.0
' Auther		wangyw@tcl.com
' Date			2016/05/17

Set WShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'set config path
configPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path & "\commitaid"
Set env = WShell.Environment("user")
If env("COMMIT_CONFIG_PATH") <> configPath Then
	MsgBox "Begin set COMMIT_CONFIG_PATH ENV. It maybe need some time."
	env.item("COMMIT_CONFIG_PATH") = configPath
End If

'need check is install or update
If objFSO.fileExists(configPath & "\config.ini") Then
	'Update
	action = "Update "
	'TODO update user config file
	
	objFSO.DeleteFile(configPath & "\config_update.ini")
	MsgBox "Update Success"
Else
	'Install
	MsgBox "Begin edit the config file"
	objFSO.MoveFile configPath & "\config_update.ini", configPath & "\config.ini"
	
	WShell.run "notepad " & configPath & "\config.ini"
End If