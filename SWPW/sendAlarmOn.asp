<%
sw=trim(request("sw"))
devN=trim(request("dev"))
dbn="./dev.mdb"
  s1=""
  s2=sw
   if file_exist(dbn) then
     	set cn=db_connection(dbn)
	 	SQL="select * from Devices where Name='"&devN&"'"
	 	
		set rs1=open_recordset(cn,sql,3,3)
'ÃÑ§O½X	Name	PubPort_In	switchCTL	Locate	PubPort_Out swState
		if not (rs1.eof and rs1.bof) then
		    ss= trim(rs1("cmd"))
		    while ss <> ""
		        ss= trim(rs1("cmd"))
		    wend
		    
		    rs1("cmd")="AlarmOn"
		    rs1.update
		    rs1.close
     		s1="OK"
        end if  
        cn.close    
   end if   

    s=s1&"|"&s2    
    response.write escape(s) 
'    response.write s
    response.end
%>    


<%
function ddformat(dd)
    yy= year(dd)
    mm= month(dd)
    if len(mm) = 1 then
        mm="0"&mm
    end if        
    tdd= day(dd)
    if len(tdd) = 1 then
        tdd="0"&tdd
    end if
    dd= yy&"/"&mm&"/"&tdd    
	ddformat=dd    
end function

function ddformatC(dd)
    yy= year(dd)
    mm= month(dd)
    if len(mm) = 1 then
        mm="0"&mm
    end if        
    tdd= day(dd)
    if len(tdd) = 1 then
        tdd="0"&tdd
    end if
    dd= yy&mm&tdd    
	ddformatC=dd    
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