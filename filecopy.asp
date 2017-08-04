<%
  dim d,date1
  dd=trim(request("d"))
  if dd="" then
	d=date()
  else
	d=cdate(left(dd,4)&"-"&mid(dd,5,2)&"-"&right(dd,2))
  end if
  
  date1=year(d) & iif (month(d)<10,"0","") & month(d) & iif(day(d)<10 ,"0","") & day(d)
  date2=year(d) & "-" & month(d) & "-" & day(d)
  date3=iif (month(d)<10,"0","") & month(d) & iif(day(d)<10 ,"0","") & day(d)
  date4=year(d) & "-" & iif (month(d)<10,"0","") & month(d) & "-" & iif(day(d)<10 ,"0","") & day(d)
  date5=year(date()) & iif (month(date())<10,"0","") & month(date()) & iif(day(date())<10 ,"0","") & day(date())
  filepath=server.mappath("./")
  
  dim source(10000,4)

  reqfunc=lcase(trim(request("reqfunc")))
  cfgfile="./" & reqfunc & ".cfg"

  Dim FSO
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")

  n=0
  if FSO.FileExists(Server.MapPath(cfgfile) ) then
    Set cfgFileObj = fso.opentextfile(server.mappath(cfgfile),1,true)
    While not cfgFileObj.AtEndOfStream
      line=trim(cfgFileObj.ReadLine)
      if len(line)>0 then
        linetext=split(line,",")
        if left(trim(linetext(0)),1) <> "#" then
		  'source(n,0),业务
			source(n,0)=trim(linetext(0))
		  'source(n,1),源路径
            linetext(1)=replace(linetext(1),"%YYYYMMDD%",date1)
            linetext(1)=replace(linetext(1),"%YYYY-MM-DD%",date4)
            linetext(1)=replace(linetext(1),"%MMDD%",date3)
            linetext(1)=replace(linetext(1),"%YYYY-M-D%",date2)
			source(n,1)=trim(linetext(1))
		  'source(n,2),源文件名
            linetext(2)=replace(linetext(2),"%YYYYMMDD%",date1)
            linetext(2)=replace(linetext(2),"%YYYY-MM-DD%",date4)
            linetext(2)=replace(linetext(2),"%MMDD%",date3)
            linetext(2)=replace(linetext(2),"%YYYY-M-D%",date2)
			source(n,2)=trim(linetext(2))
		  'source(n,3),目标路径
            linetext(3)=replace(linetext(3),"%YYYYMMDD%",date1)
            linetext(3)=replace(linetext(3),"%YYYY-MM-DD%",date4)
            linetext(3)=replace(linetext(3),"%MMDD%",date3)
            linetext(3)=replace(linetext(3),"%YYYY-M-D%",date2)
			source(n,3)=trim(linetext(3))
		  'source(n,4),false(备用)或解压密码
			source(n,4)=trim(linetext(4))

		  if instr(linetext(2),"?") >0 or  instr(linetext(2),"*") >0 then
			d0=source(n,0)
			d1=source(n,1)
			d3=source(n,3)
			d4=source(n,4)
			patrn=source(n,2)
			patrn=replace(patrn,".","\.")
			patrn=replace(patrn,"?","[\W\w]{1}")
			patrn=replace(patrn,"*","(.*?)")
			patrn="^"& patrn & "$"
		    Set Fso1=server.createobject("Scripting.FileSystemObject")
			if Fso1.folderexists(server.mappath(source(n,1))) then
				mm=0
				set mydir=Fso1.getfolder(server.mappath(source(n,1)))
'				response.write source(n,1)
				for each item in mydir.files
					Set regEx = New RegExp ' 建立正则表达式。
					regEx.Pattern = patrn ' 设置模式。
					regEx.IgnoreCase = True ' 设置是否区分大小写。
					regEx.Global = True ' 设置全程可用性。
					Set Matches = regEx.Execute(item.name) ' 执行搜索。
					if Matches.count > 0 then
						mm=1
						For Each Match in Matches
							if right(lcase(Match.Value),3)<>".ok" then
								source(n,0)=d0
								source(n,1)=d1
								source(n,2)=Match.value
								source(n,3)=d3
								source(n,4)=d4
								n=n+1
							end if
						next
					else
'						n=n+1
					end if
				next
				if mm=0 then
					n=n+1
				end if
				set mydir=nothing
			else
				n=n+1
			end if
			set Fso1=nothing
		  else
		    n=n+1
		  end if
          'n=n+1
        end if
      end if
    Wend
    cfgFileObj.Close

  end if

  req=trim(request("a"))
  action=trim(request("c"))
  if len(req)>0 and action="copy" then
    if req="all" then
      for yy=0 to n-1
		if instr(source(yy,2),"*") = 0 and  instr(source(yy,2),"?") = 0 then
			File_COPY Server.MapPath( source(yy,1) & source(yy,2)), iif( source(yy,3)="null","null",Server.MapPath( source(yy,3) & source(yy,2))), source(yy,4)
		end if
      next
    else
		if instr(source(req,2),"*") = 0 and  instr(source(req,2),"?") = 0 then
			File_COPY Server.MapPath( source(req,1) & source(req,2)), iif(source(req,3)="null","null",Server.MapPath( source(req,3) & source(req,2))), source(req,4)
		end if
    end if
  end if

  url = Request.ServerVariables("SCRIPT_NAME")
  urlParts = Split(url,"/")
  pageName = urlParts(UBound(urlParts))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<meta http-equiv="refresh" content="300; url=<%=pageName%>?reqfunc=<%=reqfunc%>"/>-->
<title>XXXXXXXX文件拷贝工具</title>
<style type="text/css">
<!--
body {
	color:#000000;
	background-color:#B3B3B3;
	margin:0;
}

#container {
	margin-left:auto;
	margin-right:auto;
	text-align:center;
	}

a img {
	border:none;
}

-->
</style>

</head>
<body>
<div id="container">
<center>
<table border=1>
  <tr>
    <form method="get" action="<%=pageName%>" name="inputdate">
    <td colspan=3>处理日期是：<input type="text" name="d" value="<%=date1%>"><input type="submit" value="改日期"><input type="hidden" name="reqfunc" value="<%=reqfunc%>">&nbsp;<input type="button" value="今天日期" onclick="todayclick('<%=date5%>')">, 
<%
    if reqfunc="qs" then
      response.write "<b>清算</b>"
    else
      if reqfunc="fdep" then
		response.write "<b>深证通</b>"
	  end if
    end if
%>
    任务数: <%=n%>, a=<%=req%>, c=<%=action%></td>
	</form>
    <td><a href="http://<%=request.ServerVariables("LOCAL_ADDR")%>/">回主页</a></td>
  </tr>
  <tr>
    <td align=center>业务</td><td align=center>源文件</td><td align=center>目标文件</td><td align=center>操作</td>
  </tr>
<%
    for xx=0 to n-1
		response.write "<tr>" & chr(13)
'	  response.write source(xx,1) & "------" & xx & "-----" & source(xx,2) & "<br>"
		source_file=Server.MapPath( source(xx,1) ) & "\" & source(xx,2)
		if lcase(source(xx,3)) <> "null" then
			target_file=Server.MapPath( source(xx,3) ) & "\" & source(xx,2)
			target_ok_file=Server.MapPath( source(xx,3) ) & "\OK-" & source(xx,2)
		else
			target_file="/*无需拷贝*/"
		end if

		source_status=FSO.FileExists( source_file )
		if source_status then
			set sf=FSO.getfile(source_file)
			source_size=sf.size
			source_time=sf.DateLastModified
		end if
		target_status=FSO.FileExists( target_file ) or FSO.FileExists( target_ok_file )
'		if target_status then
'			set tf=FSO.getfile(target_file)'
'			target_size=sf.size
'			target_time=sf.DateLastModified
'		end if
			
		
'		IF instr(lcase(source_file),"rptfile")>0 THEN
'			SH_File = Get_SH_File(source_file,source(xx,2))
'		ELSE
			SH_File = ""
'		END IF

		response.write "<td>" & source(xx,0) & "</td>"
		response.write "<td align=left width='450'>" & iif(source_status,"<font color='blue'>","<font color='red'>") & source_file & "</font>" & iif(source_status, "<font color='" & iif(source_size=0,"red","green") & "'>(" & source_size & "字节&nbsp;" & source_time & ")" & "</font>"  , "" ) & "</td>" & chr(13)
		response.write "<td align=left width='450'>" & iif(target_status,"<font color='blue'>","<font color='red'>") & target_file &  "</font><!--" & iif(target_status,"("& target_size & "字节&nbsp;" & target_time & ")" ,"" ) & "--></td>" & chr(13)
		response.write "<td>"
		IF  SH_File = "" THEN
			if (source(xx,4)="OnlyToday") and (d<>date()) then
				response.write "Out of Date" & chr(13)
			else
				if (source_status and (not target_status)) then 
					response.write "<a href='" & pageName & "?a=" & xx & "&c=copy&reqfunc=" & reqfunc & "&d="&date1
					if (source(xx,4)="OK") then
						if (lcase(source(xx,3))="null") then
							response.write "'>置OK</a>"
						else
							response.write "'>拷贝&置OK</a>"
						end if
					Else
						if target_file="/*无需拷贝*/" then
							response.write "'></a>OK"
						else
							response.write "'>拷贝</a>"
						end if
					end if
					response.write chr(13)
				else
					if not source_status then
						response.write "Checking" & chr(13)
					else
						response.write "OK" & chr(13)
					end if
				end if
			end if
		Else
			response.write SH_File & chr(13)
		END IF
		response.write "</td></tr>" & chr(13)
    next
    response.write "<tr><td align=center><a href='" & pageName & "?reqfunc=" & reqfunc & "&d=" &date1 & "'>刷新</a></td>" & chr(13)
    response.write "<td colspan=3 align=right><a href='" & pageName & "?a=all&c=copy&reqfunc=" & reqfunc & "&d=" & date1 & "'>全部拷贝或置OK</a></td></tr>" & chr(13)
%>
</table>
</center>
</div>
</body>
<script>
function todayclick(datetodayclick)
{ 
 this.inputdate.d.value=datetodayclick;
 this.inputdate.submit();
} 
</script> 
</html>
<%
  Set objFSO = Nothing
%>
<%
Function IIf(bExp1, sVal1, sVal2)
    If (bExp1) Then
        IIf = sVal1
    Else
        IIf = sVal2
    End If
End Function

Function File_COPY(src,tag,pwd)
	Dim FSO1
	Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
'    response.write "src=" & src & ", tag=" & tag & ", pwd=" & pwd & "<br/>"
  
	IF (pwd="OnlyToday") and (d<>date()) then
		'跳过严格当天拷贝文件
	else
		if FSO1.FileExists( src ) then
			if left(lcase(tag),4) <> "null" then
				tag_path=mid(tag,1,instrrev(tag,"\"))
				if not FSO.FolderExists(tag_path) then
				   FSO.CreateFolder(tag_path)
				end if
				FSO1.CopyFile src,tag,true
			end if

			if pwd="OK" then
				src1=Server.MapPath("\") & "\OK.txt"
				if left(lcase(tag),4)="null" then
					tag1=src & ".OK"
				else
					tag1= tag & ".OK"
				end if
		'		response.write src1 & "->" & tag1
				set okfile=FSO1.createtextfile(tag1,true)
				okfile.write ""
				okfile.close
				set okfile=nothing
				
			end if
			if (right(lcase(src),4)=".zip" or right(lcase(src),4)=".rar") and (right(lcase(tag),4)<> "null" or pwd="decomp") then
				if right(lcase(src),4)=".rar" then
					command= chr(34)& "c:\program files\winrar\rar.exe" & chr(34) &" x -r -o+ -ilog "
				end if
				if right(lcase(src),4)=".zip" then 
					command= chr(34)& "c:\program files\winrar\winrar.exe" & chr(34) &" x -r -o+ -ilog "
				end if
				if pwd <> "" and pwd <> "false" and pwd <> "OnlyToday" and pwd <> "decomp" then
					command = command & " -p" & pwd & " "
				else
					command = command & " -p- "
				end if
				if pwd<>"decomp" then
					command = command & tag & " " & tag_path
				else
					command = command & src & " " & mid(src,1,instrrev(src,"\"))
				end if
			  Set WshShell = server.CreateObject("Wscript.Shell")
		'	  command = command & chr(34)
		''	  response.write command
			  IsSuccess = WshShell.Run (command, 1, true)
			  response.write IsSuccess
		'	  FSO1.DeleteFile tag,true
		'	  response.write tag_path & "OK-" & mid(tag,instrrev(tag,"\")+1)
			  FSO1.CopyFile tag, tag_path & "OK-" & mid(tag,instrrev(tag,"\")+1) ,true
			  FSO1.DeleteFile tag, true
			  Set WshShell = Nothing
			end if

		end if
	end if
	Set FSO1 = Nothing
End Function

Function Get_SH_File(src,srcfile)
  Dim FSO1
  Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
  EzTrans_status_File=left(src,instr(lcase(src),"rptfile")+7)&"EzTrans_status.txt" 
'  Get_SH_File=EzTrans_status_File
  if FSO1.FileExists(EzTrans_status_File) then
	set status=FSO1.opentextfile(EzTrans_status_File,1,false)
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
  Set FSO1 = Nothing
End Function
%>
