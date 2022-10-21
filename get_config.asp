<%
  Server.ScriptTimeout=999
  dim startime, startime1, startime2, startime3
  startime=timer()

  ''Response.Buffer = False
  dim d,date1
  dd=trim(request("d"))
  if dd="" or len(dd)<>8 then
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
  
  dim source(10000,5)
  dim lines(1000)

  reqfunc=lcase(trim(request("reqfunc")))
  cfgfile="./" & reqfunc & ".cfg"

  Dim FSO
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")

  n=0
  l=0
  if FSO.FileExists(Server.MapPath(cfgfile) ) then
    Set cfgFileObj = fso.opentextfile(server.mappath(cfgfile),1,true)
    While not cfgFileObj.AtEndOfStream
      line=trim(cfgFileObj.ReadLine)
      if len(line)>0 and left(line,1) <> "#" then
		line=replace(line,"%YYYYMMDD%",date1)
		line=replace(line,"%YYYY-MM-DD%",date4)
		line=replace(line,"%MMDD%",date3)
		line=replace(line,"%YYYY-M-D%",date2)
		lines(l)=line
		l = l + 1
	  end if
	wend
	cfgFileObj.Close
	startime1=timer()
	for each line in lines
        linetext=split(line,",")
        if ubound(linetext) > 3 then			'缺少字段的不处理
			source(n,0)=trim(linetext(0))		'source(n,0),业务
			source(n,1)=server.mappath(trim(linetext(1)))		'source(n,1),源路径
			source(n,2)=trim(linetext(2))		'source(n,2),源文件名
			source(n,3)=server.mappath(trim(linetext(3)))		'source(n,3),目标路径
			source(n,4)=trim(linetext(4))		'source(n,4),false(备用)或解压密码
			source(n,5)=""					'source(n,5),一部分已预读文件信息: (日期时间:文件长度)
			
			if source(n,4) = "zxjt_gzb" then
				'zzzzzzzzzzz特别处理'
				d0=source(n,0)
				d1=source(n,1)
				d2=source(n,2)
				d3=source(n,3)
				d4=source(n,4)
				gzb_files = zxjt_gzb_getfile(source(n,1))
				if len(gzb_files) > 0 then
					q = 0
					for each gzb in split(gzb_files, ",")
''						response.write n
						if len(trim(gzb)) > 0 then
							n = n + q
							source(n,0)=d0
							source(n,1)=gzb
							source(n,2)=""
							source(n,3)=d3 & d2
							source(n,4)=d4
							q = 1
						end if
					next
				else
					'未找到zzzzzzzzzz文件'
				end if
			end if
			if instr(linetext(2),"?") >0 or instr(linetext(2),"*") >0 then
				'有通配符的特殊处理
				d0=source(n,0)
				d1=source(n,1)
				d2=source(n,2)
				d3=source(n,3)
				d4=source(n,4)
''				patrn=source(n,2)
''				patrn=replace(patrn,".","\.")
''				patrn=replace(patrn,"?","[\W\w]{1}")
''				patrn=replace(patrn,"*","(.*?)")
''				patrn="^"& patrn & "$"

				if Fso.folderexists(d1) then
					mm=0
''					Set regEx = New RegExp ' 建立正则表达式。
''					regEx.IgnoreCase = True ' 设置是否区分大小写。
''					regEx.Global = True ' 设置全程可用性。
''					regEx.Pattern = patrn ' 设置模式。
''					set mydir=Fso.getfolder(source(n,1))
'					response.write source(n,1)
''					for each item in mydir.files
''	
''						if regEx.Test(item.name) then
''							mm=1
''							if right(lcase(item.name),3) <> ".ok" then
''								source(n,0)=d0
''								source(n,1)=d1
''								source(n,2)=item.name
''								source(n,3)=d3
''								source(n,4)=d4
''								n=n+1
''							end if
''						end if
''					next
''					set mydir=nothing
''					set regEx=nothing
					for each item in split( Get_folder_files(d1 & "\" & d2) ,",")
''						response.write "item = " & item & "</br>"
						mm=mm+1
						source(n,0)=d0  & "_(" & mm & ")"
						source(n,1)=d1
						source(n,2)=split(item,"|")(0)
						source(n,3)=d3
						source(n,4)=d4
						source(n,5)=split(item,"|")(1)
						n=n+1
					next
					if mm=0 then
						n=n+1
					end if
				else
					n=n+1
				end if
			else
				n=n+1
			end if
        end if
    next
  end if
  startime2=timer()
'  target="\Server\qs\" & date1 & "\"
  req=trim(request("a"))
  action=trim(request("c"))
  if len(req)>0 and action="copy" then
    if req="all" then
      for yy=0 to n-1
		if instr(source(yy,2),"*") = 0 and  instr(source(yy,2),"?") = 0 and source(yy,4)<>"ctod" then
			File_COPY source(yy,1) & source(yy,2), iif( source(yy,3)="null","null" , source(yy,3) & source(yy,2)), source(yy,4)
		end if
      next
    else
		if instr(source(req,2),"*") = 0 and  instr(source(req,2),"?") = 0 then
			File_COPY source(req,1) & source(req,2), iif(source(req,3)="null","null", source(req,3) & source(req,2)), source(req,4)
		end if
    end if
  end if

  url = Request.ServerVariables("SCRIPT_NAME")
  urlParts = Split(url,"/")
  pageName = urlParts(UBound(urlParts))
  startime3=timer()
  ''Response.Flush
%>
