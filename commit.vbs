# $language = "VBScript"
# $interface = "1.0"
' Version		1.3
' Auther		wangyw@tcl.com
' Date			2015/11/03

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim TortoiseSVNPath  
Dim WsitaWorkPath    
Dim LsitaWorkPath    
Dim cpp_switch		 
Dim css_switch		 
Dim js_switch		 
Dim html_switch		 
Dim h_switch		 
Dim c_switch		 
Dim cmake_switch	 
Dim modify_in_days

Function GetProfile(strFileName, strSection, strName)     
	Dim st, idx    
	st = False    
	Const ForReading = 1     
	Set eXmlFile = objFSO.OpenTextFile(strFileName, ForReading)  
	 
	Do While not eXmlFile.AtEndOfStream     
		eStrLine = eXmlFile.Readline 
		If Trim(eStrLine) <> "" Then  
			If Left(Trim(eStrLine), 1) = "[" And Trim(Mid(Trim(eStrLine), 2, Len(Trim(eStrLine)) - 2)) = strSection Then     
				st = True     
			End If  
			If st = True Then
				tmp = split(Trim(eStrLine), "=")    
				If Trim(tmp(0)) = strName Then    
					GetProfile = Trim(tmp(1))   
					eXmlFile.Close     
					Set eXmlFile = Nothing     
					Exit Function  
				End If
			End If
		End If       
	Loop   
	eXmlFile.Close     
	Set eXmlFile = Nothing 
	GetProfile = ""
End Function  

Function getpathname(filename)
	
	Dim startindex, endindex

	If cpp_switch And (InStr(filename,".cpp ") <> 0) Then
		endindex = InStr(filename,".cpp") + 4
	ElseIF css_switch And (InStr(filename,".css ") <> 0) Then
		endindex = InStr(filename,".css") + 4
	ElseIF cmake_switch And (InStr(filename,".cmake ") <> 0) Then
		endindex = InStr(filename,".cmake") + 6
	ElseIF js_switch And (InStr(filename,".js ") <> 0) Then
		endindex = InStr(filename,".js") + 3
	ElseIF html_switch And (InStr(filename,".html ") <> 0) Then
		endindex = InStr(filename,".html") + 5
	ElseIF h_switch And (InStr(filename,".h ") <> 0) Then
		endindex = InStr(filename,".h") + 2
	ElseIF c_switch And (InStr(filename,".c ") <> 0) Then
		endindex = InStr(filename,".c") + 2
	Else
		endindex = 0
	End If
	
	startindex = InStr(filename,"tbrowser")
	
	If startindex <> 0 And endindex <> 0 Then
		getpathname = Mid(filename, startindex, endindex - startindex)
	Else
		getpathname = "NOT C FILE"
	End If
End Function


Function getfilelist
	Dim filecnt, filelist(1000)
	Dim cfilecnt, cfilelist(1000)
	Dim index, cfilebuf
	Dim isExists, modifyDate, nowDate

	filecnt = 0
	Do 
		filelist(filecnt) = crt.screen.Get(crt.screen.CurrentRow - 1 - filecnt, 0, crt.screen.CurrentRow - 1 - filecnt, crt.screen.Columns)
		filecnt = filecnt + 1
	Loop While InStr(filelist(filecnt-1), "$") = 0
	filecnt = filecnt - 1
	
	cfilecnt = 0
	nowDate = now()
	For index = 0 To filecnt-1
		cfilebuf = getpathname(filelist(index))
		If StrComp(cfilebuf, "NOT C FILE") <> 0 Then
			cfilebuf = Replace (cfilebuf, "/", "\")
			cfilebuf = WsitaWorkPath & cfilebuf
			isExists = objFSO.fileExists(cfilebuf)
			If isExists Then
				Set fn = objFSO.GetFile(cfilebuf)
				modifyDate = fn.DateLastModified
				If (modify_in_days = 0) Or (DateDiff("d", modifyDate, nowDate) <= modify_in_days) Then
					cfilelist(cfilecnt) = cfilebuf
					cfilecnt = cfilecnt + 1
				End If
				Set fn = Nothing
			End If
		End If
	Next
	
	If cfilecnt = 0 Then
		getfilelist = "NO FILE"
	Else
		getfilelist = """"
		For index = 0 To cfilecnt-1
			getfilelist = getfilelist & cfilelist(index)
			If index <> cfilecnt-1 Then
			getfilelist = getfilelist & "*"
			End if
		Next
		getfilelist = getfilelist & """"
	End If	
End Function

Function getConfig(filename)
	
	Dim curPath, isExists
	curPath = objFSO.GetFolder(".").Path
	curPath = curPath & "\" & filename
	isExists = objFSO.fileExists(curPath)
	If Not isExists Then
		MsgBox "No find config file"
		getConfig = 0
		Exit Function
	End If
	
	TortoiseSVNPath = GetProfile(curPath, "system", "TortoiseSVNPath")
	WsitaWorkPath = GetProfile(curPath, "system", "WsitaWorkPath")
	LsitaWorkPath = GetProfile(curPath, "system", "LsitaWorkPath")
	cpp_switch =  GetProfile(curPath, "fileswitch", "cpp_switch")
	css_switch = GetProfile(curPath, "fileswitch", "css_switch")	 
	js_switch  = GetProfile(curPath, "fileswitch", "js_switch")
	html_switch = GetProfile(curPath, "fileswitch", "html_switch")
	h_switch = GetProfile(curPath, "fileswitch", "h_switch")
	c_switch = GetProfile(curPath, "fileswitch", "c_switch")
	cmake_switch = GetProfile(curPath, "fileswitch", "cmake_switch")
	modify_in_days = GetProfile(curPath, "modifytime", "modify_in_days")
	
	getConfig = 1
End Function

Sub Main
	If getConfig("config.ini") = 0 Then
		Exit Sub
	End If
	
	crt.screen.Send "cd " & LsitaWorkPath & chr(13)
	crt.screen.WaitForString  "$"
	crt.screen.Send "svn st -q" & chr(13)
	crt.screen.WaitForString  "$"
	
	Dim commitFilePath
	Dim cmd
	commitFilePath = getfilelist()
	If commitFilePath = "NO FILE" Then
		MsgBox "no file need commit"
	Else
		cmd = """"
		cmd = cmd & TortoiseSVNPath
		cmd = cmd & "bin\TortoiseProc.exe"
		cmd = cmd & """"
		cmd = cmd & " /command:commit /path:" &commitFilePath
		cmd = cmd & " /closeonend"
		set a=createobject("wscript.shell")
		a.run cmd
	End IF
End Sub