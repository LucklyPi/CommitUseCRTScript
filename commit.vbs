# $language = "VBScript"
# $interface = "1.0"
' Version		1.4
' Auther		wangyw@tcl.com
' Date			2016/03/03

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim TortoiseSVNPath		'TortoiseSVN安装目录  
Dim WsitaWorkPath    
Dim LsitaWorkPath

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


Function getConfig()
	
	'获取配置文件路径，实际在crt中执行过程中会出现找不到配置文件的情况，
	'怀疑是因为objFSO.GetFolder(".").Path有问题，这时重新选择脚本执行即可。
	'未找到解决的办法，暂时引入环境变量COMMIT_CONFIG_PATH明确指出配置文件路径。
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
	
	'获取配置信息
	TortoiseSVNPath = GetConfigValue(configFilePath, "system", "TortoiseSVNPath")
	WsitaWorkPath   = GetConfigValue(configFilePath, "system", "WsitaWorkPath") 
	LsitaWorkPath	= GetConfigValue(configFilePath, "system", "LsitaWorkPath")
	If TortoiseSVNPath = "" Or WsitaWorkPath = "" Or LsitaWorkPath = "" Then
		MsgBox "system config ERROR"
		getConfig = 0
		Exit Function
	End If
	
	CommitFileType = MySplit(GetConfigValue(configFilePath, "commit", "CommitFileType"), ";") 
	NotCommitFile = MySplit(GetConfigValue(configFilePath, "commit", "NotCommitFile"), ";")
	ModifyInDays	= GetConfigValue(configFilePath, "commit", "ModifyInDays")
	
	getConfig = 1
	
End Function

Function CheckFile(winFileName)
	Dim pass

	'提交文件存在检查
	isExists = objFSO.fileExists(winFileName)
	If NOT isExists Then
		CheckFile = 0
		Exit Function
	End If
	
	'提交类型检查
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
	
	'提交文件黑名单检查
	For i = 0 To UBound(NotCommitFile)
		If NotCommitFile(i) <> "" AND InStr(winFileName, NotCommitFile(i)) <> 0 Then
			pass = 0
			Exit For
		End If
	Next
	If pass <> 1 Then
		CheckFile = 0
		Exit Function
	End If
	
	'提交文件修改日期限制检查
	nowDate = now()
	Set fn = objFSO.GetFile(winFileName)
	modifyDate = fn.DateLastModified
	Set fn = Nothing
	If (ModifyInDays = 0) Or (ModifyInDays - DateDiff("d", modifyDate, nowDate) > 0) Then
		CheckFile = 1
	Else
		CheckFile = 0
	End IF
End Function


Function getCommitFileList(cmdReturnStr)

	fileList = split(cmdReturnStr, vbCrLf)
	Dim commitFileList()
	Dim commitFileCnt
	Redim commitFileList(UBound(fileList)+1)
	commitFileCnt = 0
	'第一行和最后一行非文件信息，不处理
	For i = 1 To UBound(fileList)-1
		'前8个字符是svn的状态字符，暂时处理
		If Len(fileList(i)) > 9 Then
			linuxFileName = Mid(fileList(i), 9, Len(fileList(i)) - 8)
			winFileName   = WsitaWorkPath & Replace(LinuxFileName, "/", "\")
			If CheckFile(winFileName) Then
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
	crt.screen.Send "cd " & LsitaWorkPath & chr(13)
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