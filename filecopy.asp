<!-- #include file="get_config.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<meta http-equiv="refresh" content="300; url=<%=pageName%>?reqfunc=<%=reqfunc%>"/>-->
<title>wowowowowo文件拷贝工具</title>
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
table{
	margin:0px auto;
	font:Georgia 11px;
	color:#333333;
	text-align:center;
	border-collapse:collapse;
}
td{
	border:1px solid #000;
}
#loading { 
  position:absolute; 
  width:124px; 
  height:124px; 
  top:50%;
  left:50%; 
  margin: -62px -62px;
  background-color:#FFFFFF;
  border:1px solid #CCCCCC;
  text-align:center;
  padding:20px;
}
</style>

</head>
<body>
<div id="loading" style="display:none">
	<img src="loading.gif" /> Loading...
</div>
<div id="container">
<center>
<form method="get" action="<%=pageName%>" name="inputdate">
<table border=1>
  <tr>  
    <td colspan=3>处理日期是：<input type="text" name="d" value="<%=date1%>"><input type="submit" value="改日期" onclick="document.getElementById('loading').style.display='';"><input type="hidden" name="reqfunc" value="<%=reqfunc%>">&nbsp;<input type="button" value="今天日期" onclick="todayclick('<%=date5%>')">, 
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
    <td><a href="http://<%=request.ServerVariables("LOCAL_ADDR")%>/">回主页</a></td>
  </tr>
  <tr>
    <td align=center>业务</td><td align=center>源文件</td><td align=center>目标文件</td><td align=center>操作</td>
  </tr>
<%
	On Error Resume Next
    for xx=0 to n-1
		response.write "<tr>" & chr(13)
'	  response.write source(xx,1) & "------" & xx & "-----" & source(xx,2) & "<br>"
		source_file=source(xx,1) & "\" & source(xx,2)
		If Err.Number = 0 Then
			source_status=FSO.FileExists( source_file )
		else
			source_file=source(xx,1)
			source_status=false
		end if

		if lcase(source(xx,3)) <> "null" then
			target_file= source(xx,3) & "\" & source(xx,2)
			target_ok_file=source(xx,3) & "\OK-" & source(xx,2)
		else
			target_file="/*无需拷贝*/"
			target_ok_file=""
		end if
		source_size=0
		source_time=""
		if source_status then
			if source(xx,5) = "" then
				set sf=FSO.getfile(source_file)
				source_size=sf.size
				source_time=sf.DateLastModified
			else
				source_time=split(source(xx,5),"%")(0)
				source_size=split(source(xx,5),"%")(1)
			end if
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
		response.write "<td align=left width='450'>" & iif(source_status,"<font color='blue' size=-1>","<font color='red' size=-1>") & source_file & "</font>" & iif(source_status, "<font color='" & iif(source_size=0,"red","green") & "' size=-1>(" & source_size & "字节&nbsp;" & source_time & ")" & "</font>"  , "" ) & "</td>" & chr(13)
		response.write "<td align=left width='450'>" & iif(target_status,"<font color='blue' size=-1>","<font color='red' size=-1>") & target_file &  "</font><!--" & iif(target_status,"("& target_size & "字节&nbsp;" & target_time & ")" ,"" ) & "--></td>" & chr(13)
		response.write "<td>"
		IF  SH_File = "" THEN
			if (source(xx,4)="OnlyToday") and (d<>date()) then
				response.write "Out of Date" & chr(13)
			else
				if (source_status and (not target_status)) then 
					response.write "<a href='#' onclick="&chr(34)&"document.getElementById('loading').style.display='';window.location.href='" & pageName & "?a=" & xx & "&c=copy&reqfunc=" & reqfunc & "&d="&date1
					if (source(xx,4)="OK") then
						if (lcase(source(xx,3))="null") then
							response.write "';" & chr(34) & ">置OK</a>"
						else
							response.write "';" & chr(34) & ">拷贝&置OK</a>"
						end if
					elseif (source(xx,4)="ctod") then
						response.write "';" & chr(34) & ">处理</a>"
					else
						if target_file="/*无需拷贝*/" then
							response.write "';" & chr(34) & "></a>OK"
						else
							response.write "';" & chr(34) & ">拷贝</a>"
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
		''Response.Flush
    next
    response.write "<tr><td align=center>"
''	response.write "<a href='" & pageName & "?reqfunc=" & reqfunc & "&d=" &date1 & "' onclick=" & chr(34) & "document.getElementById('loading').style.display='';" & chr(34) &">刷新</a>"
	response.write "<a href='#' onclick=" & chr(34) & "document.getElementById('loading').style.display='';window.location.href='" & pageName & "?reqfunc=" & reqfunc & "&d=" &date1 & "';" & chr(34) & ">刷新</a>"
	response.write "</td>" & chr(13) & "<td align=left>用时(秒): "
	response.write "读配置:" & int(startime1-startime) & "&nbsp;查文件:" & int(startime2-startime1) & "&nbsp;,拷文件" & int(startime3-startime2) & "&nbsp;,显文件" & int(timer()-startime3)
    response.write "</td>" & chr(13) & "<td colspan=2 align=right>"
''	response.write "<a href='" & pageName & "?a=all&c=copy&reqfunc=" & reqfunc & "&d=" & date1 & "' onclick=" & chr(34) & "document.getElementById('loading').style.display='';" & chr(34) & ">全部拷贝或置OK</a>"
	response.write "<a href='#' onclick=" & chr(34) & "document.getElementById('loading').style.display='';window.location.href='" & pageName & "?a=all&c=copy&reqfunc=" & reqfunc & "&d=" &date1 & "';" & chr(34) & ">全部拷贝或置OK</a>"
	response.write "</td></tr>" & chr(13)
	On Error GoTo 0
%>
</table>
</form>
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

<!-- #include file="functions.asp" -->
