<%
Function IIf(bExp1, sVal1, sVal2)
    If (bExp1) Then
        IIf = sVal1
    Else
        IIf = sVal2
    End If
End Function

Function CreateFolders(path)
	''Set fso = CreateObject("Scripting.FileSystemObject")
	CreateFolderEx fso,path
	''Set fso = Nothing
End Function

Function CreateFolderEx(fso, path)
	If fso.FolderExists(path) Then
		Exit Function
	End if
	If Not fso.FolderExists(fso.GetParentfolderName(path)) Then
		CreateFolderEx fso, fso.GetParentfolderName(path)
	End IF
	fso.CreateFolder(path)
End Function

Function File_COPY(src, tag, pwd)
	''Dim FSO1
	''Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
'    response.write "src=" & src & ", tag=" & tag & ", pwd=" & pwd & "<br/>"
  
	IF (pwd="OnlyToday") and (d<>date()) then
		'跳过严格当天拷贝文件
	else
		if FSO.FileExists( src ) then
			if left(lcase(tag),4) <> "null" then
				tag_path=mid(tag,1,instrrev(tag,"\"))
				tag_file=mid(tag,instrrev(tag,"\")+1)
				''response.write "tag_path=" & tag_path
				if not FSO.FolderExists(tag_path) then
					CreateFolders tag_path
				   ''FSO1.CreateFolder(tag_path)
				end if
				if pwd="decomp" then
					tag_file = "OK-" + tag_file
					tag = tag_path + tag_file
				end if
				FSO.CopyFile src,tag,true
				if not FSO.FileExists(tag) then
					response.write "Copy Fail : " & tag
				end if
			end if

			if pwd="OK" then
		'		src1=Server.MapPath("\") & "\OK.txt"
				if left(lcase(tag),4)="null" then
					tag1 = src & ".OK"
				else
					tag1 = tag & ".OK"
				end if
		'		response.write src1 & "->" & tag1
				set okfile=FSO.createtextfile(tag1,true)
				okfile.write ""
				okfile.close
				set okfile=nothing
				
			end if
			if (right(lcase(src),4)=".zip" or right(lcase(src),4)=".rar") and (right(lcase(tag),4)<> "null" or pwd="decomp" or pwd="nofolder") and ( pwd <> "OK" and pwd <> "false" ) then
				if right(lcase(src),4)=".rar" then
					command= chr(34)& "c:\program files\winrar\rar.exe" & chr(34) 
				end if
				if right(lcase(src),4)=".zip" then 
					command= chr(34)& "c:\program files\winrar\winrar.exe" & chr(34)
				end if
				if pwd="nofolder" then
					command = command & " e"
				else
					command = command & " x"
				end if
				command = command & " -r -o+ -dh -inul -ilogd:\rar.log "
				if pwd <> "" and pwd <> "false" and pwd <> "OnlyToday" and pwd <> "decomp" and pwd <> "nofolder" then
					command = command & " -p" & pwd & " "
				else
					command = command & " -p- "
				end if
				if pwd<>"decomp" then
					command = command & tag & " " & tag_path
				else
					if tag="null" then
						command = command & src & " " & mid(src,1,instrrev(src,"\"))
					else
						command = command & tag & " " & tag_path
					end if
				end if
				Set WshShell = server.CreateObject("Wscript.Shell")
		'		command = command & chr(34)
''				IsSuccess = 999
				IsSuccess = WshShell.Run (command, 1, True)
''				do while IsSuccess = 999
''					response.write "wait for winrar running decompress"
''					call DelayTime(1)
''				loop 
''				if IsSuccess=0 then
''					FSO1.CopyFile tag, tag_path & "OK-" & mid(tag,instrrev(tag,"\")+1) ,true
''					FSO1.DeleteFile tag, true					
'				else
'					response.write "decomp error: " & cstr(IsSuccess) & " ,Command=" &command & "<br/>"
'					response.write IsSuccess
''				end if
				Set WshShell = Nothing
			end if
			if pwd="ctod" then
				'对yyyyyyyyy字段特殊处理'
				src_path=mid(src,1,instrrev(src,"\")-1)
				src_file=mid(src,instrrev(src,"\")+1)
				format_modify src_path, src_file
			end if
			if pwd="zxjt_gzb" then
				'对zzzzzzzz特殊处理'
				zxjt_gzb_special src, tag, d
			end if
		end if
	end if
	''Set FSO1 = Nothing
End Function

Function Get_SH_File(src,srcfile)
  ''Dim FSO1
  ''Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
  EzTrans_status_File=left(src,instr(lcase(src),"rptfile")+7)&"EzTrans_status.txt" 
'  Get_SH_File=EzTrans_status_File
  if FSO.FileExists(EzTrans_status_File) then
	set status=FSO.opentextfile(EzTrans_status_File,1,false)
	do while status.AtEndOfStream = false
		getaline=status.ReadLine
		if left(getaline,9) = "File = " & ucase(left(srcfile,2)) then
			for nn = 1 to 10
				status.skipline
			next
			getaline=status.readline
			Get_SH_File=trim(mid(getaline,9))
		end if
	loop
	status=close
	set status=Nothing
  end if
  ''Set FSO1 = Nothing
End Function

Function format_modify(path, src)
	Dim FSO1,conn,connstr,rs
	Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
	On Error Resume Next 
	if FSO1.FileExists(  path & "\" & src )then
''		Set conn = CreateObject("ADODB.Connection")
''		connstr = "Driver={Microsoft dBASE VFP Driver (*.dbf)};SourceType=DBF;SourceDB=" & Server.MapPath(path) & ";Exclusive=No"
''		conn.Open connstr
''		sql="select * from " & src
''		set rs=createobject("adodb.recordset")
''		set rs=conn.Execute(sql)
''		rs.open sql,conn,0,3
''		dzrq_old=rs("DZRQ")
''		rs.Close()
''		set rs=nothing
''		conn.Close()
''		set conn=nothing
		'先备份源文件'
		FSO1.CopyFile path & "\" & src, path & "\" & src & ".ZSZQ", False
		''
		temp=split(src,"_")
		dzrq_old=replace(temp(2),".dbf","")
		dzrq_new=mid(dzrq_old,5,2) & "/" & mid(dzrq_old, 7) & "/" & mid(dzrq_old,3,2)
''		TimeDelaySeconds(5)
		connstr = "Driver={Microsoft dBASE VFP Driver (*.dbf)};SourceType=DBF;SourceDB=" & path & ";Exclusive=Yes"
		Set conn = CreateObject("ADODB.Connection")
		conn.Open connstr
		sql="Alter Table " & src & " Alter COLUMN DZRQ Date"
		conn.Execute sql
		sql="update " & src & " set DZRQ=ctod('" & dzrq_new & "')"
		conn.Execute sql
		conn.Close
		set conn=Nothing
''		response.write "done"
	else
''		response.write path & "\" & src & " not found"
		set FSO1=Nothing
	end if
	On Error GoTo 0
End Function

Function zxjt_gzb_special(src_file, tag_file, tag_date)
''	response.write "gzb:<br/>" & src_file & "<br/>" & tag_file & "<br/>" & d & "<br/>"

	Dim FSO1, conn, connstr, rs
	Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
	On Error Resume Next 
	if FSO1.FileExists(  src_file )then
		path = left(tag_file, instrrev(tag_file, "\") - 1 )
		src = replace(mid(tag_file, instrrev(tag_file, "\") + 1), ".DBF", "")
		newdate=split(d, "/")
		Set conn = CreateObject("ADODB.Connection")
		connstr = "Driver={Microsoft dBASE VFP Driver (*.dbf)};SourceType=DBF;SourceDB=" & path & ";Exclusive=No"
		conn.Open connstr

		sql = "update " & src & " set Ffdate=ctod('" & newdate(1) & "/" & newdate(2) & "/" & right(newdate(0),2) & "')"
		conn.Execute sql
		conn.Close
		set conn=Nothing

	else
		response.write src_file & " not found"
	end if
	set FSO1=Nothing
	On Error GoTo 0
End Function

Function zxjt_gzb_getfile(src_file)
	result = ""
	temp=split(src_file,"*")
	''Set Fso1=server.createobject("Scripting.FileSystemObject")
	if Fso.folderexists(temp(0)) then
		mm=0
		set mydir=Fso.getfolder(temp(0))
		for each item in mydir.SubFolders
			gzb_path = replace(src_file , "*" , item.name)
			if Fso.FileExists(gzb_path) then
				set sf=FSO.getfile(gzb_path)
				if sf.size > 1024 then
					result = gzb_path & "," & result
				end if
			end if
		next
	else
		''response.write "zzzzzzzz源路径未找到"
	end if
	zxjt_gzb_getfile = result
	set mydir = nothing
	set sf = nothing
	''set Fso1 = nothing
End Function

Sub DelayTime(secondNumber) 
	dim startTime 
	startTime=NOW() 
	do while datediff("s",startTime,NOW())<secondNumber 
	loop 
End Sub 

Function Get_folder_files(fs_path)
	Set WshShell = server.CreateObject("Wscript.Shell")
    Set IsSuccess = WshShell.exec ("%windir%\system32\cmd.exe /s /c " & chr(34) & "dir /a:-d /-C " & fs_path & " 2>&1" & chr(34))
    result=IsSuccess.stdout.readall()
    Set IsSuccess = Nothing
    Set WshShell = Nothing
	Get_folder_files = ""
	for each line in split(result, chr(10))
		if left(line,1)="2" then
			Get_folder_files = Get_folder_files & "," & replace(mid(line,37),chr(13),"") & "|" & left(line,17) & "%" & mid(line,19,17)
		end if
	next
	Get_folder_files = mid(Get_folder_files,2)
''	response.write "return=" & Get_folder_files & "</br>"
End Function

Set FSO = Nothing

%>
