# $language = "VBScript"
# $interface = "1.0"
' Version		1.4
' Auther		wangyw@tcl.com
' Date			2016/05/16

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim TortoiseSVNPath		'TortoiseSVN安装目录  
Dim WWorkPath    
Dim LWorkPath

Dim CommitFileType
Dim NotCommitFile
Dim ModifyInDays

Function GetConfigValue(configFilePath, section, name)     
	Dim inSection
	inSection = 0
	
	Set configFile = objFSO.OpenTextFile(configFilePath, 1)
	Do While not configFile.AtEndOfStream     
		strLine = Trim(configFile.Readline)
		If strLine <> "" Then 
			If Left(strLine, 1) = "[" Then
				If Trim(Mid(strLine, 2, Len(strLine)-2)) = section Then 
					inSection = 1
				Else
					inSection = 0
				End If
			End if
			If inSection = 1 AND InStr(strLine, "=") <> 0 Then
				tmp = split(strLine, "=")   
				If Trim(tmp(0)) = name Then    
					GetConfigValue = Trim(tmp(1))   
					configFile.Close     
					Set configFile = Nothing     
					Exit Function
				End If				
			End If
		End If       
	Loop   
	configFile.Close     
	Set configFile = Nothing 
	GetConfigValue = ""
End Function  

Function MySplit(expression, delimiter)
	expression_Trim = Trim(expression)
	target = split(expression_Trim, delimiter)
	target_Ubound = Ubound(target)
	counter = 0
	Dim target_Trim()
	redim target_Trim(target_Ubound+1)
	For i = 0 To target_Ubound 
		tmp = Trim(target(i))
		IF tmp <> "" Then
			target_Trim(counter) = tmp
			counter = counter+1
		End If
	next
	ReDim Preserve target_Trim(counter)
	MySplit = target_Trim
End Function

Function FormatPath(path, os)
    path = Trim(path)
	If os = "win" Then
		If Right(path, 1) <> "\" Then
			path = path & "\"
		End If
    Else
		If Right(path, 1) <> "/" Then
			path = path & "/"
		End If
	End If
	FormatPath = path
End Function


Function getConfig()
	
	Dim configFilePath
	Set env = CreateObject("WScript.Shell").Environment("user")
	configFilePath = env("COMMIT_CONFIG_PATH")
	If Len(configFilePath) = 0 Then
		Set env = CreateObject("WScript.Shell").Environment("system")
		configFilePath = env("COMMIT_CONFIG_PATH")
	End If
	If Len(configFilePath) = 0 Then
		configFilePath = objFSO.GetFolder(".").Path
	End If
	configFilePath = configFilePath & "\" & "config.ini"
	If Not objFSO.fileExists(configFilePath) Then
		MsgBox "No find config file: " & configFilePath
		getConfig = 0
		Exit Function
	End If
	
	TortoiseSVNPath = FormatPath(GetConfigValue(configFilePath, "system", "TortoiseSVNPath"), "win")
	WWorkPath       = FormatPath(GetConfigValue(configFilePath, "system", "WsitaWorkPath"), "win")
	LWorkPath	    = FormatPath(GetConfigValue(configFilePath, "system", "LsitaWorkPath"), "linux")
	If TortoiseSVNPath = "" Or WWorkPath = "" Or LWorkPath = "" Then
		MsgBox "system config ERROR"
		getConfig = 0
		Exit Function
	End If
	
	CommitFileType = MySplit(GetConfigValue(configFilePath, "commit", "CommitFileType"), ";") 
	NotCommitFile = MySplit(GetConfigValue(configFilePath, "commit", "NotCommitFile"), ";")
	ModifyInDays	= GetConfigValue(configFilePath, "commit", "ModifyInDays")
	
	getConfig = 1
	
End Function

Function CheckFile(winFileName, fileStatus)
	Dim existCheck
	Dim typeCheck
	Dim blackCheck
	Dim timeCheck
	existCheck = 1
	typeCheck = 1
	blackCheck = 1
	timeCheck = 1
	
	'set check item
	If Left(fileStatus, 1) = "D" Then
		'delete file
		existCheck = 0
		timeCheck = 0
	End If
	If Mid(fileStatus, 2, 1) = "D" Then
		'delete dir
		existCheck = 0
		typeCheck = 0
		timeCheck = 0
	End If
	
	'begin file check
	'file exist check
	If existCheck <> 0 Then
		isExists = objFSO.fileExists(winFileName)
		If NOT isExists Then
			CheckFile = 0
			Exit Function
		End If
	End if
	
	'file type check
	If typeCheck <> 0 Then
		pass = 0
		tmp = split(winFileName, ".")
		fileType = "." & tmp(UBound(tmp))
		For i = 0 To UBound(CommitFileType)
			If fileType = CommitFileType(i) Then
				pass = 1
				Exit For
			End If
		Next
		If pass <> 1 Then
			CheckFile = 0
			Exit Function
		End If
	End if
	
	If blackCheck <> 0 Then
		For i = 0 To UBound(NotCommitFile)
			If NotCommitFile(i) <> "" AND InStr(winFileName, NotCommitFile(i)) <> 0 Then
				CheckFile = 0
				Exit Function
			End If
		Next
	End If
	
	If timeCheck <> 0 And existCheck <> 0 Then
		nowDate = now()
		Set fn = objFSO.GetFile(winFileName)
		modifyDate = fn.DateLastModified
		Set fn = Nothing
		If (ModifyInDays = 0) Or (ModifyInDays - DateDiff("d", modifyDate, nowDate) > 0) Then
		Else
			CheckFile = 0
			Exit Function
		End IF
	End IF
	CheckFile = 1
End Function


Function getCommitFileList(cmdReturnStr)
	fileList = split(cmdReturnStr, vbCrLf)
	Dim commitFileList()
	Dim commitFileCnt
	Redim commitFileList(UBound(fileList)+1)
	commitFileCnt = 0
	For i = 1 To UBound(fileList)-1
		If Len(fileList(i)) > 9 Then
			fileStatus    = Left(fileList(i), 9)
			linuxFileName = Mid(fileList(i), 9, Len(fileList(i)) - 8)
			winFileName   = WWorkPath & Replace(LinuxFileName, "/", "\")
			If CheckFile(winFileName, fileStatus) Then
				commitFileList(commitFileCnt) = winFileName
				commitFileCnt = commitFileCnt + 1
			End If
		End If
	Next
	
	If commitFileCnt = 0 Then
		getCommitFileList = "NO FILE"
	Else
		getCommitFileList = """"
		For index = 0 To commitFileCnt-1
			getCommitFileList = getCommitFileList & commitFileList(index)
			If index <> commitFileCnt-1 Then
			getCommitFileList = getCommitFileList & "*"
			End if
		Next
		getCommitFileList = getCommitFileList & """"
	End If 
End Function




Sub Main
	If getConfig() = 0 Then
		Exit Sub
	End If
	crt.screen.IgnoreEscape = False
	crt.screen.Send "cd " & LWorkPath & chr(13)
	crt.screen.WaitForString  "$"
	crt.screen.Send "svn st -q" & chr(13)
	receive = crt.Screen.ReadString("$", 60)
	
	Dim commitFileList
	Dim cmd
	commitFileList = getCommitFileList(receive)
	If commitFileList = "NO FILE" Then
		MsgBox "no file need commit"
	Else
		cmd = """"
		cmd = cmd & TortoiseSVNPath
		cmd = cmd & "bin\TortoiseProc.exe"
		cmd = cmd & """"
		cmd = cmd & " /command:commit /path:" &commitFileList
		cmd = cmd & " /closeonend"
		set a=createobject("wscript.shell")
		a.run cmd
	End IF
End Sub