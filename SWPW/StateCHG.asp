<%
db=trim(request("db"))
devN=trim(request("dev"))
dt=trim(request("dt"))
sst=trim(request("sst"))
ft=trim(request("ft"))
if dt <> "" then
	idt= Cint(dt)
	if idt >=1 and idt <= 23 then
	    idt=idt
	else
	    idt=24
	end if
else
	idt=4
end if

s=TCtime(sst,idt)
	
ta=split(s,"|")

s1=""
s2=""
sum=0.0
k=0.0
dim temp
for i=0 to ubound(ta)
    tb= split(ta(i),":")
    dd=tb(0)
    tt=tb(1)
    tc=split(tt,";")
    if ubound(tc)> 0 then
    	tt=tc(0)
    	et=tc(1)
    else
        tt=tc(0)
        et=""
    end if
    	
    ss=getCurrentRec(db,dd,tt,et,k)

    tc=split(ss,"|")
    if s1= "" then
        s1=tc(0)
        s2=tc(1)
        sum=CDbl(tc(2))
    else
        s1=s1&","&tc(0)
        s2=s2&","&tc(1)
        sum=sum+ CDbl(tc(2))
        
    end if
 

next     

    if k=0.0 then
        sum=0
    else    
        sum= ((sum/k)*(idt)*110/ 1000 )  /1000 
    end if
    
	dbn="./dev.mdb"

   if file_exist(dbn) then
     	set cn=db_connection(dbn)
	 	SQL="select * from Devices where Name='"&devN&"'"
	 	
		set rs1=open_recordset(cn,sql,3,3)
'ÃÑ§O½X	Name	PubPort_In	switchCTL	Locate	PubPort_Out swState
		if not (rs1.eof and rs1.bof) then
		    sw=rs1("swState")

		    rs1.close
        end if  
        cn.close   
         
   		cmd=""
		if ft <> "" then
	    	ta=split(ft,":")
	    	if temp < ta(0) then
	        	if sw = "false" then
	            	cmd="on"
	        	end if
	    	end if
	    	if temp > ta(1) then
	        	if sw="true" then
	            	cmd="off"
	        	end if
	    	end if
	    	
	    	if cmd <> "" then
     			set cn=db_connection(dbn)
	 			SQL="select * from Devices where Name='"&devN&"'"
	 	
				set rs1=open_recordset(cn,sql,3,3)
'ÃÑ§O½X	Name	PubPort_In	switchCTL	Locate	PubPort_Out swState
				if not (rs1.eof and rs1.bof) then
		    		rs1("cmd")=cmd
		    		rs1.update
		    		rs1.close
        		end if  
        		cn.close    
	        end if
		end if    
   
	end if
    s=s1&"|"&s2&"|"&sw&"|"& digit2(sum)&"|"&temp
    response.write escape(s) 
'    response.write s
    response.end
%>    


<%

function getCurrentRec(db,dd,tt,tte,k)
	dim dbn,cn,sql,rs,s1,s2,sum,ssum,st,et

	dbn="./"&db&"/"&dd&".mdb"
    ssum=" tt"
	s1=""
	s2=""  
	sum=0.0
	et= trim(tte)

   if file_exist(dbn) then
   

     set cn=db_connection(dbn)
     if et="" then
     	SQL="select * from DATA where Time >='"&tt&"' order by time"  
     else
     	SQL="select * from DATA where Time >='"&tt&"' and Time <= '"&et&"' order by time"  
     end if
     
 '    SQL="select * from DATA "  
 
     set rs=open_recordset(cn,sql,3,3)
     if not(rs.eof and rs.bof) then
        while not rs.eof 
     		if s1="" then
         		s1= CTime(trim(rs("Time")))
         		s2= trim(rs("current"))
         		sum= CDbl( s2)
     		else
         		s1= s1&","&CTime(trim(rs("Time")))

         		if trim(rs("current")&"  ") = "" then
         		    f=CDbl(st)
         		    s2= s2&","&st
         		else
         		    f= CDbl(rs("current"))
         		    s2= s2&","&trim(rs("current"))
         		    st=trim(rs("current"))
         		end if
         		
         		sum= sum + f
    		end if 
    		temp= trim(rs("Temp"))
    		rs.movenext
    		k=k+1
    	wend   
     	rs.close     
     end if  
     cn.close    
   end if  
 
   ssum=sum
   getCurrentRec=s1&"|"&s2&"|"&ssum
end function

function digit2(ff)
dim ta,ffs

    ff=ff+0.005
    ffs=cstr(ff)
    if instr(ffs,"-") then
    	ta=split(ffs,"e")
    	if Cint(ta(1)) > 2 then
    		ffs="0.00"
    	else
    	    ffs="0.01"
    	end if
    else  
    	ta=split(ffs,".")
    	if ubound(ta) > 0 then
        	ffs= ta(0)&"."&mid(ta(1),1,2)
    	else
        	ffs= ta(0)
    	end if
	end if
    digit2=ffs    
end function

function TCtime(sst,n)
	dim ta,tb,H,hs,mss,ms,sss,ss,dd,tt,taa,tag,hsN	
	if sst="" then
		dd=date   
		tt=now
		taa= split(tt," ")
		tt= taa(2)
		tag= taa(1)	       		
	else
	    ta=split(sst," ")	    
	    
	    if ubound(ta)= 0 then
	        dd=date
	        tt= now
			taa= split(tt," ")
			tt= taa(2)
			tag= taa(1)	        
	    else
	        dd= Cdate(trim(ta(0)))
	    	tt= trim(ta(1))
	    	tag= ""
	    end if	    
	end if
	
	ddf=ddformat(dd)
	ddt= ddformatC(dd)
	
	tb=split(tt,":")
	
	H= int(tb(0))
    
		if tag="¤U¤È" then
	    	if H >= 12 then
	        	H=H-12
	    	end if
			H= H+12
		end if
	if H < 10 then
    	hsN="0"&H
    else
        hsN=H
	end if

		if H >= n then
    		H= H-n
    		dd1=""
    		H1=0
		else
	    	dd1=DateAdd("d", -1, dd)
	    	ddf1=ddformat(dd1)
	    	ddt1= ddformatC(dd1)
	    	H1= 24-(n-H)
	    
			H=0
		end if

	hs=H
	if H < 10 then
    	hs="0"&H
	end if
	hs1=H1
	if H1 < 10 then
    	hs1="0"&Cstr(H1)
	end if
	
	mss=""
	ms= int(tb(1))
	mss=ms
	if ms < 10 then
    	mss="0"&ms
	end if
	sss=""
	ss= int(tb(2))
	sss=ss
	if ss < 10 then
    	sss="0"&ss
	end if

	
	if dd1 = "" then
	    TCtime=ddt&":"&ddf&"/"&hs&"/"&mss&"/"&sss&";"&ddf&"/"&hsN&"/"&mss&"/"&sss
	else
	    TCtime=ddt1&":"&ddf1&"/"&hs1&"/"&mss&"/"&sss&"|"&ddt&":"&ddf&"/00/00/00"&";"&ddf&"/"&hsN&"/"&mss&"/"&sss
    end if	    
  
end function

function ddformat(dd)
dim ddd
    yy= year(dd)
    mm= month(dd)
    if len(mm) = 1 then
        mm="0"&mm
    end if        
    tdd= day(dd)
    if len(tdd) = 1 then
        tdd="0"&tdd
    end if
    ddd= yy&"/"&mm&"/"&tdd    
	ddformat=ddd    
end function

function ddformatC(dd)
dim ddd
    yy= year(dd)
    mm= month(dd)
    if len(mm) = 1 then
        mm="0"&mm
    end if        
    tdd= day(dd)
    if len(tdd) = 1 then
        tdd="0"&tdd
    end if
    ddd= yy&mm&tdd    
	ddformatC=ddd    
end function

function CTime(s)
    dim ta,dd
    ta=split(s,"/")
    dd= ta(0)&"/"&ta(1)&"/"&ta(2)&" "&ta(3)&":"&ta(4)&":"&ta(5)    
	CTime=dd    
end function

function file_exist(tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   tfln1= server.mappath(tfln)
'   tfln1= tfln
   
   if fso.FileExists(tfln1) then
       file_exist = true
   else
       file_exist =false  
   end if     
end function

SUB WriteAllTextFile(sfln,s)
  Const ForReading = 1, ForWriting = 2
  Dim fso, f
  file_delete(sfln)
  Set fso = CreateObject("Scripting.FileSystemObject")
  absfln=server.mappath(sfln)
  Set f = fso.OpenTextFile(absfln, ForWriting, True)
  f.WriteLine s
  f.Close
End SUB
sub file_delete(tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
'   tfln1= server.mappath(tfln)
   tfln1=tfln
   if fso.FileExists(tfln1) then
       Set MyFile = fso.GetFile(tfln1)

       MyFile.Delete  
   end if     
end sub

Function ReadAllTextFile(sfln)
on error resume next
  Const ForReading = 1, ForWriting = 2
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  absfln=server.mappath(sfln)
'  absfln=sfln
  Set f = fso.OpenTextFile(absfln, ForReading)
  ReadAllTextFile =  f.ReadAll
  if err.number <> 0 then
      response.write "["&sfln&"]["&absfln&"]<br>"
      response.end
  end if
End Function

function db_connection(dbn)
dim provider,db_path,ds,cns,conn

'provider = "Provider = Microsoft.Jet.OLEDB.4.0; "
provider = "Provider =Microsoft.ACE.OLEDB.12.0; "
db_path= server.mappath(dbn)
'db_path= dbn
ds= "Data Source= "& db_path
cns= provider&ds

set conn= server.createobject("ADODB.Connection")
conn.open cns
set db_connection = conn
end function

function open_recordset(cnn,sql,cs,ltype)
dim rs
set rs= server.createobject("ADODB.Recordset")
rs.open sql,cnn,cs,ltype
set open_recordset = rs
end function

function SQLdb_connection(dbsv,uid,pwd,dbn)

dim provider,ds,usr,pswd,dbfile,cns,conn

provider = "Provider = SQLOLEDB.1; "
ds= "Data Source= "& dbsv&"; "
usr= "User ID="&uid&"; "
pswd= "Password="&pwd&";"
dbfile= "Initial Catalog="&dbn

cns= provider&ds&usr&pswd&dbfile

set conn= server.createobject("ADODB.Connection")
conn.open cns

set SQLdb_connection = conn

end function

function GetMdbRecordset(dbn, SQL)
dim cn,rs

set cn=db_connection(dbn)
set rs=open_recordset(cn,SQL,2,3)
set GetMdbRecordset=rs
end function

%>