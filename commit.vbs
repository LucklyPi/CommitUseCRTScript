# $language = "VBScript"
# $interface = "1.0"
' Version		1.2
' Auther		wangyw@tcl.com
' Date			2015/11/02



'系统环境配置
TortoiseSVNPath  = "C:\Program Files\TortoiseSVN\"    'TortoiseSVN安装目录
WsitaWorkPath    = "Z:\sita\"						  'windows下工程目录
LsitaWorkPath    = "~/samba/sita/"					  'Linux下工程目录

'提交文件配置
cpp_switch		 = 1 	'.cpp文件提交开关
css_switch		 = 1 	'.css文件提交开关
js_switch		 = 1	'.js文件提交开关
html_switch		 = 1 	'.html文件提交开关
h_switch		 = 1 	'.h文件提交开关
c_switch		 = 1 	'.c文件提交开关
cmake_switch	 = 0    '.cmake文件提交开关

modify_in_days	 = 5	'在最近多少天内修改过的文件才会提交,0代表无限大



Dim cfilecnt
Dim cfilelist(1000)

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
	Dim filecnt
	Dim filelist(1000)
	
	filecnt = 0
	Do 
		filelist(filecnt) = crt.screen.Get(crt.screen.CurrentRow - 1 - filecnt, 0, crt.screen.CurrentRow - 1 - filecnt, crt.screen.Columns)
		filecnt = filecnt + 1
	Loop While InStr(filelist(filecnt-1), "$") = 0
	filecnt = filecnt - 1
	
	Dim index
	Dim cfilebuf
	cfilecnt = 0
	For index = 0 To filecnt-1
		cfilebuf = getpathname(filelist(index))
		If StrComp(cfilebuf, "NOT C FILE") <> 0 Then
			cfilelist(cfilecnt) = cfilebuf
			cfilecnt = cfilecnt + 1
		End If
	Next
	
	For index = 0 To cfilecnt-1
		cfilelist(index) = Replace (cfilelist(index), "/", "\")
		cfilelist(index) = WsitaWorkPath & cfilelist(index)
	Next
	
	
End Function


Sub Main	
	crt.screen.Send "cd " & LsitaWorkPath & chr(13)
	crt.screen.WaitForString  "$"
	crt.screen.Send "svn st -q" & chr(13)
	crt.screen.WaitForString  "$"
	getfilelist

	Dim cmd
	Dim nowDate
	Dim modifyDate
	Set fso=createobject("Scripting.FileSystemObject")
	nowDate = now()
	If cfilecnt = 0 Then
		MsgBox "no file need commit"
	Else
		cmd = """"
		cmd = cmd & TortoiseSVNPath
		cmd = cmd & "bin\TortoiseProc.exe"
		cmd = cmd & """"
		cmd = cmd & " /command:commit /path:"
		cmd = cmd & """"
		Dim index
		For index = 0 To cfilecnt-1
			Set fn=fso.GetFile(cfilelist(index))
			modifyDate = fn.DateLastModified
			If (modify_in_days = 0) Or (DateDiff("d", modifyDate, nowDate) <= modify_in_days) Then
				cmd = cmd & cfilelist(index)
				If index <> cfilecnt-1 Then
				cmd = cmd & "*"
				End if
			End If
		Next
		cmd = cmd & """"
		cmd = cmd & " /closeonend"
		set a=createobject("wscript.shell")
	a.run cmd
	End IF
End Sub