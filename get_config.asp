<%
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
        if left(trim(linetext(0)),1) <> "#" and ubound(linetext) > 3 then
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
			if source(n,4) = "zxjt_gzb" then
				'zzzzz估值表特别处理'
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
					'未找到zzzzz估值表文件'
				end if
			end if
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
'  target="\Server\qs\" & date1 & "\"
  req=trim(request("a"))
  action=trim(request("c"))
  if len(req)>0 and action="copy" then
    if req="all" then
      for yy=0 to n-1
		if instr(source(yy,2),"*") = 0 and  instr(source(yy,2),"?") = 0 and source(yy,4)<>"ctod" then
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
