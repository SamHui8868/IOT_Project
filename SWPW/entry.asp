<%
session.TimeOut=120
devN=trim(request("devN"))
if devN="" then
	devN="test11"
end if

dbn="dev.mdb"
set cn1=db_connection(dbn)
SQL="select * from Devices where Name='"&devN&"'"
set rs1=open_recordset(cn1,sql,3,3)
'識別碼	Name	PubPort_In	switchCTL	Locate	PubPort_Out swState

db=rs1("PubPort_In")
Locate=rs1("Locate")
swState=rs1("swState")
rs1.close
cn1.close

tt=now
ta=split(tt," ")
tt= ta(2)
tb=split(tt,":")
H= int(tb(0))
if ta(1)="下午" then
	H= H+12
end if
if H > 4 then
    H= H-4
else
	H=0
end if
hs=H
if H < 10 then
    hs="0"&H
end if
mss=""
ms= int(tb(1))
mss=ms
if ms < 10 then
    mss="0"&ms
end if
sss=""
ss= int(tb(1))
sss=ss
if ss < 10 then
    sss="0"&ss
end if
    
dd=date   

tt=ddformat(dd)
tt=tt&"/"&hs&"/"&mss&"/"&sss
dd= ddformatC(dd)

s1=""
s2=""

'dd="20200505"
dbn="./"&db&"/"&dd&".mdb"   
   if file_exist(dbn) then
     set cn=db_connection(dbn)
          
     SQL="select * from DATA where Time >='"&tt&"'"
     set rs=open_recordset(cn,sql,3,3)
     if not(rs.eof and rs.bof) then
        k=0
        while not rs.eof 
     		if s1="" then
         		s1= CTime(trim(rs("Time")))
         		st=s1
         		s2= trim(rs("current"))
     		else
     	 		st=CTime(trim(rs("Time")))
         		s1= s1&","&st
         		s2= s2&","&trim(rs("current"))
    	 	end if 
    	 	rs.movenext
    	 	k=k+1
    	wend   
     	rs.close     
     end if  
     cn.close    
   end if   
   
dbn="devActive.mdb"   
Stime=""
Etime=""
WD =""
   if file_exist(dbn) then
     set cn=db_connection(dbn)
     SQL="select * from ActTable where Name='"&devN&"'"
     set rs=open_recordset(cn,sql,3,3)
     if not(rs.eof and rs.bof) then

        while not rs.eof 
     	if Stime="" then
         Stime= trim(rs("Stime"))
         Etime= trim(rs("Etime"))
         WD= trim(rs("WeekDay"))
     	else
         Stime= Stime&"|"&trim(rs("Stime"))
         Etime= Etime&"|"&trim(rs("Etime"))
         WD= WD&"|"&trim(rs("WeekDay"))
    	end if 
    	rs.movenext
    	k=k+1
    	wend   
     	rs.close     
     end if  
     cn.close    
   end if   
Stime=Stime&"|"
Etime=Etime&"|"
WD=WD&"|"

%>
<html>
<head>
  <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
    
google.charts.load('current', {'packages':['line']});

<!--    
//==========================================================
var ttcount=0;
function createXHR(){

  if(window.XMLHttpRequest){
    xmlHttp = new XMLHttpRequest;
  }else if (window.ActiveXObject){
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  }
  
  if(!xmlHttp){
    alert('您使用的瀏覽器不支援Ajax');
    return false;
  }
}
//======================================
function StateCHG(db)
{

  var ss,n,sn,k;
  
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
    createXHR();
    
   var dt=f1.T5.value;
   var sst=f1.SStime.value;
   var ft;
   
    ft=f1.T8.value+":"+f1.T9.value;

   	if(!f1.C1.checked){
   		ft="";
   	}
  
   var url = 'StateCHG.asp?n='+ds+' &db='+db+'&dev=<%=devN%>'+'&dt='+dt+'&sst='+sst+'&ft='+ft;
    
 //   location.href=url;
    
    window.status = 'State Change Command...';
    xmlHttp.open('post',url,true);
    xmlHttp.onreadystatechange = catchStateCHG;
    xmlHttp.send(null);
    ttcount++;
 
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}
function catchStateCHG()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        if(ss != ""){
        	ss = unescape(ss);

            ta= ss.split("|"); 
            
        	f1.T1.value=ta[0];
            f1.T2.value = ta[1];
//            f1.B2.innerText=ta[2];
            f1.B2.value=ta[2];
            document.getElementById('cdegree').innerHTML = ta[3];
            f1.T6.value=ta[4];
        }
        drawChart();
        window.status = 'Done';
    }
    
}

function swClick()
{

  var ss,n,sn,k;
    
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
    
    
    ss=f1.B2.value;

    f1.B2.disabled=true;
    sws="on";
    if(ss=="true"){
        sws="off";
    }

    if(confirm("確定要將開關["+sws+"]嗎?")==false){
    	f1.B2.disabled=false;
      
	}else{
	    createXHR();

    	var url = 'sendSWCtl.asp?n='+ds+' &dev=<%=devN%> &sw='+sws;
    
 //   location.href=url;
    
    	window.status = 'State Change SW Command...';
    	xmlHttp.open('post',url,true);
    	xmlHttp.onreadystatechange = catchswClick;
    	xmlHttp.send(null);
    	ttcount++;
 	}
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}

function catchswClick()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        if(ss != ""){
        	ss = unescape(ss);

            ta= ss.split("|"); 
            f1.B2.value="false";
//            f1.B2.innerText="false";
            if( ta[1]=="on"){	
                f1.B2.value="true";
//                f1.B2.innerText="true";
                f1.B2.value="true";
            }
        }
        
		f1.B2.disabled=false;
		
        window.status = 'Done';
    }
    
}

function HTBChange()
{

  var ss,n,sn,k;
    
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
    
    
    ss=document.f1.T7.value;
    

    if(ss.trim() != ""){

	    createXHR();

    	var url = 'HTBChange.asp?n='+ds+' &dev=<%=devN%> &HTB='+ss;
    
//    location.href=url;
    
    	window.status = 'State Change HTBChange Command...';
    	xmlHttp.open('post',url,true);
    	xmlHttp.onreadystatechange = catchHTBChange;
    	xmlHttp.send(null);
    	ttcount++;
 	}
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}
function catchHTBChange()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        
		AlarmOnClick();
        window.status = 'Done';
    }
    
}

function AlarmOnClick()
{

  var ss,n,sn,k;
    
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
 

	    createXHR();

    	var url = 'sendAlarmOn.asp?n='+ds+' &dev=<%=devN%>' ;
    
//    location.href=url;
    
    	window.status = 'State Change AlarmOff Command...';
    	xmlHttp.open('post',url,true);
    	xmlHttp.onreadystatechange = catchAlarmOnClick;
    	xmlHttp.send(null);
    	ttcount++;
 	
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}

function catchAlarmOnClick()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        
        window.status = 'Done';
    }
    
}


function AlarmOffClick()
{

  var ss,n,sn,k;
    
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
 

	    createXHR();

    	var url = 'sendAlarmOff.asp?n='+ds+' &dev=<%=devN%>' ;
    
//    location.href=url;
    
    	window.status = 'State Change AlarmOff Command...';
    	xmlHttp.open('post',url,true);
    	xmlHttp.onreadystatechange = catchAlarmOffClick;
    	xmlHttp.send(null);
    	ttcount++;
 	
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}
function catchAlarmOffClick()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        
		document.f1.T7.value="";
        window.status = 'Done';
    }
    
}

function setActime()
{

  var s1,n,sn,k,s2,s3,i,ss1,ss2,ss3;
  
  var currentTime = new Date()
  var hh = currentTime.getHours()
  var mm = currentTime.getMinutes()
  var ss = currentTime.getSeconds()
//================

    ds=hh*3+mm*7+ss*11+ttcount;
    createXHR();
    s1="";
    s2="";
    s3="";
    for(i=0;i<document.f1.Stime.length;i++){
    	ss1 = document.f1.Stime[i].value;
    	ss2 = document.f1.Etime[i].value;
    	ss3 = document.f1.WeekDay[i].value;

    	ss1=ss1.trim();
    	ss2=ss2.trim();
    	ss3=ss3.trim();
    	
    	if(ss1 != ""  && ss2!=""){
    	    if(s1==""){
    	        s1=ss1;
    	        s2=ss2;
    	        s3=ss3;
    	    }else{
    	        s1=s1+"|"+ss1;
    	        s2=s2+"|"+ss2;
    	        s3=s3+"|"+ss3;
    	    }
    	}     
    	
    }


    var url = 'setActime.asp?n='+ds+' &dev=<%=devN%> &st='+escape(s1)+'&et='+escape(s2)+'&wd='+escape(s3);
   
//    location.href=url;
    
    window.status = 'State Change SW Command...';
    xmlHttp.open('post',url,true);
    xmlHttp.onreadystatechange = catchsetActime;
    xmlHttp.send(null);
    ttcount++;
 
 //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}
function catchsetActime()
{
	var ss,ta,i,ssA,s,idn;
	ss="kkk";
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200){
      ss = xmlHttp.responseText;
    }
    if(ss != "kkk"){
        if(ss != ""){
        	ss = unescape(ss);

			document.getElementById('setInfo').innerHTML = ss;
         }
        
        window.status = 'Done';
    }
    
}



//===================================

function drawChart() 
{
    var x;
    var ss1= f1.T1.value;
    var ss2= f1.T2.value;
    
    var ta1= ss1.split(",");
    var ta2= ss2.split(",");
    
	var s=[];

	for (x = 0; x < ta1.length; x++){

 		s.push([
   			ta1[x],
   			eval(ta2[x])
 		]);
	}
	f1.T4.value= ta1[x-1];
	if(ta1[x-1].trim()==""){
    	f1.T4.value= ta1[x-2];
    }
    var data = new google.visualization.DataTable();
   
    data.addColumn('string', 'Time');
    data.addColumn('number', 'Guardians of the Current');
//      data.addColumn('number', 'The Avengers');
//      data.addColumn('number', 'Transformers: Age of Extinction');

    data.addRows(s);

    var options = {
        chart: {
          title: '<%=Locate%> ',
          subtitle: 'in milliampere (mA)'
        },
        width: 800,
        height: 300,
        axes: {
          x: {
            0: {side: 'top'}
          }
        }
    };

    var chart = new google.charts.Line(document.getElementById('line_top_x'));

    chart.draw(data, google.charts.Line.convertOptions(options));
}

function addnewRec()
{

	var s1=document.getElementById('RecList').innerHTML;
	
	var ta=s1.split("<tr>");
	var i,ss,n,s10,s11,Stime,Etime,WD,tb,tc,td,s1,s2,s3;
    s1="";
    s2="";
    s3="";
    for(i=0;i<document.f1.Stime.length;i++){
    	ss1 = document.f1.Stime[i].value;
    	ss2 = document.f1.Etime[i].value;
    	ss3 = document.f1.WeekDay[i].value;

    	ss1=ss1.trim();
    	ss2=ss2.trim();
    	ss3=ss3.trim();
    	
    	    if(s1==""){
    	        s1=ss1;
    	        s2=ss2;
    	        s3=ss3;
    	    }else{
    	        s1=s1+"|"+ss1;
    	        s2=s2+"|"+ss2;
    	        s3=s3+"|"+ss3;
    	    }
   
    	
    }

	
	ss="";
	n=ta.length;
	for(i=0;i< n-2;i++)
	    ss=ss+"<ta>"+ta[i];
	    
	ss = ss+"<ta>"+ta[i]+"<ta>"+ta[i]+"<ta>"+ta[n-1];

	document.getElementById('RecList').innerHTML= ss;

	s1=s1+"|";
	s2=s2+"|";
	s3=s3+"|";

	tb=s1.split("|");
	tc=s2.split("|");
	td=s3.split("|");
	
	for(i=0;i<tb.length;i++)
	{
		f1.Stime[i].value=tb[i];
		f1.Etime[i].value=tc[i];
		f1.WeekDay[i].value=td[i];
	}
  
}
function delRec()
{

	var s1=document.getElementById('RecList').innerHTML;
	
	var ta=s1.split("<tr>");
	var i,ss,n,s10,s11,Stime,Etime,WD,tb,tc,td,s1,s2,s3;
    s1="";
    s2="";
    s3="";
    for(i=0;i<document.f1.Stime.length;i++){
    	ss1 = document.f1.Stime[i].value;
    	ss2 = document.f1.Etime[i].value;
    	ss3 = document.f1.WeekDay[i].value;

    	ss1=ss1.trim();
    	ss2=ss2.trim();
    	ss3=ss3.trim();
    	
    	    if(s1==""){
    	        s1=ss1;
    	        s2=ss2;
    	        s3=ss3;
    	    }else{
    	        s1=s1+"|"+ss1;
    	        s2=s2+"|"+ss2;
    	        s3=s3+"|"+ss3;
    	    }
   
    	
    }
	
	ss="";
	n=ta.length;
	if(n <= 4) return;
	for(i=0;i< n-2;i++)
	    ss=ss+"<ta>"+ta[i];
	    
	ss = ss+"<ta>"+ta[n-1];

	document.getElementById('RecList').innerHTML= ss;
	
	s1=s1+"|";
	s2=s2+"|";
	s3=s3+"|";

	tb=s1.split("|");
	tc=s2.split("|");
	td=s3.split("|");
	
	for(i=0;i<n-1;i++)
	{
		f1.Stime[i].value=tb[i];
		f1.Etime[i].value=tc[i];
		f1.WeekDay[i].value=td[i];
	}

}

var count=0;
function Reloadwindow()
{
	var s= f1.T5.value;
	var n=1;
	if(s.trim()!= ""){
		n= s / 4;
		n=(n < 1)?1:n;
	}
	document.getElementById('setInfo').innerHTML =""
	StateCHG('<%=db%>');
	window.setTimeout("Reloadwindow();",15000 * n);
	count++;
}

function FTClick()
{
	var ft1= f1.T8.value;
	var ft2=f1.T9.value;
	var flag=1;
	
	if(!f1.C1.checked){
	    
	}else{
		f1.C1.checked=false;
		if(ft1.trim() == ""){
			flag=0;
		}
		if(ft2.trim() == ""){
			flag=0;
		}

		if(flag==1){
	    	if(!f1.C1.checked){
	        	f1.C1.checked=true;
	        	StateCHG('<%=db%>');
	    	}
		}
	}
}



function init()
{

	StateCHG('<%=db%>');

    window.setTimeout("Reloadwindow();",15000);
}
    
//-->    
    
//    google.charts.setOnLoadCallback(drawChart);
  </script>
</head>
<body>

  <div id="line_top_x"></div>
  
	<form method="POST" action="--WEBBOT-SELF--" name="f1">
	    　<p>  &nbsp;&nbsp;&nbsp;  電源開關:<input type="button" name="B2" value="<%=swState%>" style="width: 50px; height: 23px" onclick="swClick();">
		<input type="text" name="T4" size="20" value="<%=st%>">
		<span id="cdegree"> 
		</span>(度)
		<br><br>
		<input type="button" value="重劃圖" name="B1" onclick="init();"><input type="text" name="SStime" size="25">
		<input type="hidden" name="T1" size="20" value="<%=s1%>">
		<input type="hidden" name="T2" size="20" value="<%=s2%>">
		<input type="hidden" name="T3" size="20" value="<%=dbn%>">
		&nbsp;&nbsp;&nbsp;  :
		<input type="text" name="T5" size="4">(1~23) &nbsp;&nbsp;</p>
		<p>  室溫: <input type="text" name="T6" size="4">&nbsp;&nbsp;&nbsp;&nbsp; 高溫警示溫度:&nbsp; 
		<input type="text" name="T7" size="7" onchange="HTBChange();">&nbsp;<input type="button" name="B3" value="警報解除" style="width: 70; height: 27" onclick="AlarmOffClick();">
		</p>
		<p>  &nbsp;<input type="checkbox" name="C1" value="" onclick="FTClick();" >恆溫開關: <input type="text" name="T8" size="6"> ~ <input type="text" name="T9" size="7"></p>
		<p>  　</p>
		
<%
ta=	split(Stime,"|")
tb=split(Etime,"|")
tc=split(WD,"|")
%>	
<span id="RecList"> 
		<table border="1" width="28%">
			<tr>
				<td width="152">起始時間</td>
				<td>終止時間</td>
				<td>星期</td>
				
			</tr>
			<tr>
				<td width="152"><input type="text" name="Stime" size="20" value="<%=ta(0)%>"></td>
				<td><input type="text" name="Etime" size="20" value="<%=tb(0)%>"></td>
				<td><input type="text" name="WeekDay" size="20" value="<%=tc(0)%>"></td>
			</tr>
			
<%
for i=1 to ubound(ta)
%>			
			<tr>
				<td width="152"><input type="text" name="Stime" size="20" value="<%=ta(i)%>"></td>
				<td><input type="text" name="Etime" size="20" value="<%=tb(i)%>"></td>
				<td><input type="text" name="WeekDay" size="20" value="<%=tc(i)%>"></td>
			</tr>
<%
next
%>			
			<tr align=center>
				<td colspan="3">
				<p align="center">
				<input type="button" value=" + " name="B4" onclick="addnewRec();"><input type="button" value="送出規劃表" name="B3" onclick="setActime();">
				<input type="button" value=" - " name="B4" onclick="delRec();">
				<span id="setInfo">
				</span>	
				</p>
				
				<input type="hidden" name="Stime" size="20" value=""><input type="hidden" name="Etime" size="20" value=""><input type="hidden" name="WeekDay" size="20" value="">
				<input type="hidden" name="newRec" value="">

				</td>
			</tr>

		</table>

	</form>

</body>
</html>
<Script Language="JavaScript">

init();

</Script>

<%
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

Function GetDaysInMonth(iMonth, iYear)
	Dim dTemp
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetDaysInMonth = Day(dTemp)
End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
  Dim dTemp
  dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
  GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
  SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
  AddOneMonth = DateAdd("m", 1, dDate)
End Function

Function GetWeekStartOn(dDateInWeek)
  Dim dTemp
  dTemp = DateAdd("d", -(WeekDay(dDateInWeek) - 1), dDateInWeek)
  GetWeekStartOn = dTemp
End Function

Function WeekdayName(iWeekday)
  WeekdayName = MID("日一二三四五六週", iWeekday, 1)
End Function


sub line_Diff(kwhs,kwhs1,mdkwhs)
dim ta,tb,i,s0,s1,ss

ta=split(kwhs,"|")
tb=split(kwhs1,"|")
mdkwhs=ta(0)&"-<br>"&tb(0)
for i=1 to ubound(ta)
	ss=0.0
    s0= replace(ta(i),"KWH","")
    s1= replace(tb(i),"KWH","")
    if s0 <> 0.000 and s1 <> 0.000 then
    	ss= s0 - s1
    else
        ss=0.000
    end if
    mdkwhs=mdkwhs&"|"& Cfloat3(ss)&"KWH"
next    
end sub

sub session_DKWH(sddt, sdds, kwhs, dkwhs)
dim tkwh0,tkwh1,edd,dkwh,ta,s0,s1,sdd
dim i
tkwh0=0.1
tkwh1=0.1
sdd=sddt
ta=split(sdd,"-")
sdds="&nbsp;"
kwhs=ta(0)
dkwhs=ta(0)
for i=1 to 8

    tkwh0= getkwh(dbpath,sdd)
    edd= adddate(sdd,6)
    ta=split(sdd,"-")
    s0=ta(1)&"-"&ta(2)
    ta=split(edd,"-")
    s1=ta(1)&"-"&ta(2)
	edd= adddate(sdd,7)
	
    tkwh1= getkwh(dbpath,edd)
    if tkwh0 <> 0.000 and tkwh1 <> 0.000 then
   	dkwh0=dkwh
    	dkwh=tkwh1-tkwh0	
    else
        dkwh=0.0
        dkwh0=0.0
    end if
    
    if i=1 then
        dkwhs=dkwhs&"|"& "&nbsp;"	
    else
        if dkwh <> 0.000 and dkwh0 <> 0.000 then   
            dkwhs=dkwhs&"|"& Cfloat3(dkwh-dkwh0)&"KWH"
        else
            dkwhs=dkwhs&"|"& "0.000KWH"
        end if	
    end if
    sdds=sdds&"|"&s0&" ~ "&s1
    kwhs= kwhs&"|"& Cfloat3(dkwh)&"KWH"
    

	sdd=edd
next

end sub

function show_line(s,n,bgc)
dim ta,mn,i,ss
ta=split(s,"|")
if n= 0 then 
    mn= ubound(ta)+1
else
    mn=n
end if
show_line=mn
%>
<tr>
<%
for i= 1 to mn
%>
<%
    if i-1 <= ubound(ta) then
        ss= replace(ta(i-1),"0.000KWH","&nbsp;")
    else
        ss= "&nbsp;"
    end if
%>
<td align=center bgcolor=<%=bgc%>>
<%=ss%>
</td>  
<%               
next  
%>  
</tr>

<%
end function

function adddate(sdd,n)
dim ta,dd,mm,yy,mdn

mmi=0
yyi=0
ta=split(sdd,"-")
dd= Cint(ta(2))+n
mm= Cint(ta(1))
yy= Cint(ta(0))
mdn=GetDaysInMonth(mm, yy)
if(dd > mdn ) then
	dd=dd-mdn
    mm=mm+1
    if(mm > 12) then
        mm=1
        yy= yy+1
    end if
end if
if(dd <= 0 ) then
    mm=mm-1
    if mm = 0 then
        mm=12
        yy=yy-1
    end if
    mdn=GetDaysInMonth(mmi, yyi)
	dd=dd+mdn
end if
ta(0)=yy
ta(1)=mm
    if len(ta(1)) = 1 then
        ta(1)="0"&ta(1)
    end if        
ta(2)=dd
    if len(ta(2)) = 1 then
        ta(2)="0"&ta(2)
    end if          
adddate=ta(0)&"-"&ta(1)&"-"&ta(2)

end function

function getkwh(dbpath,sdd)
dim cn,rs,sql
dim ta
ta=split(sdd,"-")
dbn=dbpath&ta(0)&"_"&ta(1)&".mdb"
getkwh=0.0
if file_exist(dbn) then
set cn=db_connection(dbn)
sql="select * from DATA where [atw] like '"&sdd&"%' order by atw"
set rs=open_recordset(cn,sql,3,3)

if not(rs.eof and rs.bof) then
    getkwh=rs("del_total_kwh")
end if
rs.close
cn.close
else
'response.write "DB is not exist!!["&dbn&"]<br>"
end if
end function

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

sub sel_rs_option(seln,rs,jfun,skey)
dim i,opv
%>
      <select size="1" name="<%= seln %>" onchange="<%=jfun%>" id="<%= seln %>">
    <option value="">  </option>
<%

    for i=0 to rs.recordcount-1
       if rs(0) = "[其他]" then
           opv ="<-自行輸入->"
       else 
           opv = rs(0)
       end if       

       if opv = skey then
%>       
    <option selected value="<%=opv%>"><%=rs(0)%></option>
<%     else%>    
    <option  value="<%=opv%>"><%=rs(0)%></option>
<%     end if
       rs.movenext
    next
%>    
    </select>
<%   
end sub

sub sel_option(seln,idn,selA,jfun,skey)
dim i,opv
%>
      <select size="1" name="<%= seln %>" id="<%=idn%>" onchange="<%=jfun%>">
          <option  value=""></option>

<%

    for i=0 to UBound(selA)
       if selA(i)= "其他" then
           opv ="<-自行輸入->"
       else 
           opv = selA(i)
       end if       

       if opv = skey then
%>       
    <option selected value="<%=opv%>"><%=selA(i)%></option>
<%     else%>    
    <option  value="<%=opv%>"><%=selA(i)%></option>
<%     end if
    next
%>    
    </select>
</p>
<%    
end sub

function digit2(s)
dim n,ss
ss= trim(s)
if ss <> "" then
    n=len(ss)
    if n = 1 then
        ss="0"&ss
    end if
end if
digit2=ss
end function  

function set_time_stamp()
    dd= date
    tt= time
    syy= Year(dd)
    smm= Month(dd)
    if len(smm)=1 then
        smm="0"&smm
    end if    
    sdd= Day(dd)
    if len(sdd)=1 then
        sdd="0"&sdd
    end if    
    shh= Hour(tt)
    if len(shh)=1 then
        shh="0"&shh
    end if    
    
    smn= Minute(tt)
    if len(smn)=1 then
        smn="0"&smn
    end if    
    
    sss= Second(tt)
    if len(sss)=1 then
        sss="0"&sss
    end if    
    
    dds= syy&smm&sdd
    tts= shh&smn&sss

    set_time_stamp=dds&tts
end function

Function ReadAllTextFile(tfln)
	
    Const ForReading = 1, ForWriting = 2
  	Dim fso, f,tfln1
    Set fso = CreateObject("Scripting.FileSystemObject")
    tfln1= server.mappath(tfln)  
    Set f = fso.OpenTextFile(tfln1, ForReading)
    ReadAllTextFile =  f.ReadAll
End Function

sub file_move(sfln,tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   sfln1= server.mappath(sfln)
   tfln1= server.mappath(tfln)
   
   Set MyFile = fso.GetFile(sfln1)
   MyFile.Copy (tfln1)
   MyFile.Delete   
end sub

sub Response_error_Messages
 '     Response.Write("<b>Wrong selection : </b>" & Err.description&"["& mySmartUpload.Form("title").values&"]<br>")
end sub

function  get_sno()
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim fso, f,sno,isno,ssno,fln
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  fln = server.mappath("./paper_sno.txt")  
  Set f = fso.OpenTextFile(fln, ForReading)
  sno = f.ReadLine
  f.Close
 
  isno=cint(sno)+1
  ssno= snostr(isno)
  
  Set f = fso.OpenTextFile(fln, ForWriting, True)
  f.WriteLine ssno
  f.Close
  
  get_sno=ssno
  
End function

function snostr(n)

    dim s
    
    s= cstr(n)
    
    k= len(s)
    for i= k to 4
        s= "0"&s
    next   
     
    snostr= trim(s)
    
end function

function Cfloat3(nn)
dim s,ss,ni,ta,ks,i
    s= Cstr(nn)
    ta=split(s,"E-")
    
    if ubound(ta)=1 then
        k= ta(0) * 10000.0
        
        if k >= 0 then
            ks=Cstr(k)
        	for i=0 to Cint(ta(1))-2
            	ks="0"&ks
        	next
        	s="0."&ks
        else
        	k= -1 * k
        	ks=Cstr(k)
        	for i=0 to Cint(ta(1))-2
            	ks="0"&ks
        	next
        	s="-0."&ks
        end if
        
    end if
    
    ni=instr(s,".")
    if ni > 0 then
        s=s&"000"
        ss= mid(s,1,ni+3)
    else
        ss= s&".000"
    end if        
	Cfloat3=ss
end function

Function ReadAllTextFile(sfln)
  Const ForReading = 1, ForWriting = 2
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  absfln=server.mappath(sfln)
  Set f = fso.OpenTextFile(absfln, ForReading)
  ReadAllTextFile =  f.ReadAll
End Function

sub file_copy(sfln,tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   sfln1= server.mappath(sfln)
   tfln1= server.mappath(tfln)
   
   Set MyFile = fso.GetFile(sfln1)
   MyFile.Copy (tfln1)
end sub

sub file_move(sfln,tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   sfln1= server.mappath(sfln)
   tfln1= server.mappath(tfln)
   
   Set MyFile = fso.GetFile(sfln1)
   MyFile.Copy (tfln1)
   MyFile.Delete   
end sub

sub file_delete(tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   tfln1= server.mappath(tfln)
   
   if fso.FileExists(tfln1) then
       Set MyFile = fso.GetFile(tfln1)

       MyFile.Delete  
   end if     
end sub

function file_exist(tfln)
   set fso= server.createobject("Scripting.FilesystemObject")
   
   tfln1= server.mappath(tfln)
   
   if fso.FileExists(tfln1) then
       file_exist = true
   else
       file_exist =false  
   end if     
end function

function db_connection(dbn)
dim provider,db_path,ds,cns,conn

'provider = "Provider = Microsoft.Jet.OLEDB.4.0; "
provider = "Provider =Microsoft.ACE.OLEDB.12.0; "
db_path= server.mappath(dbn)
ds= "Data Source= "& db_path
cns= provider&ds

set conn= server.createobject("ADODB.Connection")
conn.open cns

set db_connection = conn

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

function open_recordset(cnn,sql,cs,ltype)
dim rs
on error resume next
set rs= server.createobject("ADODB.Recordset")
rs.open sql,cnn,cs,ltype
set open_recordset = rs
if err.number <> 0 then
response.write "["&SQL&"]"
response.end
end if

end function

function GetMdbRecordset(dbn, SQL)
dim cn,rs

set cn=db_connection(dbn)
set rs=open_recordset(cn,SQL,3,3)
set GetMdbRecordset=rs
end function
%>