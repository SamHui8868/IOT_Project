
<%


%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="zh-tw">
<TITLE>TEST01監控頁面 </TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<style type="text/css">
<!--
body {
	background-image: url('../images/line.gif');
	background-repeat: repeat-x;
	background-color: #8B816B
}
.style3 {font-size: 12px; color: #FFFFFF; }
.style5 {font-size: 12px}
.style6 {
	font-size: 12;
	color: #675C50;
}
-->
</style>
</HEAD>
<script language="JavaScript" type="text/JavaScript">
<!--
var currentTime = new Date()
var ss = currentTime.getSeconds()
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
//    document.Image51.align="right";
//    alert("yy");
    posflag=0;
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function Reloadwindow()
{
    StateCHG();
    window.setTimeout("Reloadwindow();", 15000);
}
function init()
{
    StateCHG();
    window.setTimeout("Reloadwindow();", 2000);
} 
//-->
</script>
<script type="text/javascript">
<!-- 
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
    
function StateCHG()
{

    var ss, n, sn, k;

    var currentTime = new Date()
    var hh = currentTime.getHours()
    var mm = currentTime.getMinutes()
    var ss = currentTime.getSeconds()
    //================

    ds = hh * 3 + mm * 7 + ss * 11 + ttcount;
    createXHR();

    var url = 'GetCurrent.asp?n=' + ds ;

    //   location.href=url;

    window.status = 'State Change Command...';
    xmlHttp.open('post', url, true);
    xmlHttp.onreadystatechange = catchGetCurrent;
    xmlHttp.send(null);
    ttcount++;

    //   document.getElementById('chkEmail').innerHTML = "xxx"; 

}
function catchGetCurrent() {
    var ss, ta, i, ssA, s, idn;
    ss = "kkk";
    if (xmlHttp.readyState == 4 && xmlHttp.status == 200) {
        ss = xmlHttp.responseText;
    }
    if (ss != "kkk") {
        if (ss != "") {
            ss = unescape(ss);

            ta = ss.split("|");

            document.getElementById('Current').innerHTML = ta[1];
            if (ta[2] == "true") {
                document.all('Light').src = document.MM_p[9].src;
            } else {
                document.all('Light').src = document.MM_p[8].src;
            }
        }
        window.status = 'Done';
    }

}

//-->
</script>


<script type="text/javascript">
<!-- 
function setligtOffDwn()
{
    if(document.all('BTNOFF_state').src != document.MM_p[3].src){
		document.all('BTNOFF').src=document.MM_p[3].src;
		document.all('BTNOFF_state').src=document.MM_p[3].src;
	}else{
		document.all('BTNOFF').src=document.MM_p[2].src;
		document.all('BTNOFF_state').src=document.MM_p[2].src;
	}
}
function setligtOffOv()
{
	document.all('BTNOFF').src=document.MM_p[0].src;
}
function setligtOffOut()
{
	document.all('BTNOFF').src = document.all('BTNOFF_state').src;

}

function setligtOnDwn()
{
    if(document.all('BTNON_state').src != document.MM_p[7].src){
		document.all('BTNON').src=document.MM_p[7].src;
		document.all('BTNON_state').src=document.MM_p[7].src;
	}else{
		document.all('BTNON').src=document.MM_p[4].src;
		document.all('BTNON_state').src=document.MM_p[4].src;
	}
}
function setligtOnOv()
{
	document.all('BTNON').src=document.MM_p[5].src;
}
function setligtOnOut()
{
	document.all('BTNON').src = document.all('BTNON_state').src;

}

function setBTNDwn()
{
    if(document.all('BTN_state').src == document.MM_p[4].src){
		document.all('BTN').src=document.MM_p[2].src;
		document.all('BTN_state').src=document.MM_p[2].src;
		document.all('Light').src=document.MM_p[8].src;
		
	}else{
		document.all('BTN').src=document.MM_p[4].src;
		document.all('BTN_state').src=document.MM_p[4].src;
		document.all('Light').src=document.MM_p[9].src;
	}
}
function setBTNOv()
{
	if(document.all('BTN_state').src != document.MM_p[4].src){
		document.all('BTN').src=document.MM_p[0].src;
	}else{
		document.all('BTN').src=document.MM_p[5].src;
	}
}
function setBTNOut()
{
	document.all('BTN').src = document.all('BTN_state').src;

}

function set_light() 
{

}
// -->
</script> 

<script language="JavaScript" src="../mm_menu.js"></script>

<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 onLoad="MM_preloadImages('./images/BTN_off.png','./images/off-1.png','./images/off-2.png','./images/off-3.png',
'./images/BTN_on.png','./images/on-1.png','./images/on-2.png','./images/on-3.png','./images/light_OFF.png','./images/light_ON.png');init();" >
<configuration>

    <system.web>
        <xhtmlConformance mode="Legacy" />
    </system.web>

</configuration> 

					
<a onMouseDown="setligtOffDwn()" onMouseOver="setligtOffOv()" onMouseOut="setligtOffOut()">
	<img id="BTNOFF" src="images/off-2.png" name="BTNOFF"  border="0" width="125" height="117"  alt="BTNOFF" 
	style="border: 1px ridge #0066FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
</a>

<a onMouseDown="setligtOnDwn()" onMouseOver="setligtOnOv()" onMouseOut="setligtOnOut()">
	<img id="BTNON" src="images/BTN_on.png" name="BTNON"  border="1" width="125" height="117"  alt="BTNON" 
	style="border: 1px ridge #0066FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
</a>

<br>

<a onMouseDown="setBTNDwn()" onMouseOver="setBTNOv()" onMouseOut="setBTNOut()">
	<img id="BTN" src="images/off-2.png" name="BTN"  border="1" width="125" height="117"  alt="BTN" 
	style="border: 1px ridge #0066FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
</a>

<img id="BTNOFF_state" src="./images/off-2.png" name="BTNOFF_state" width="5" height="5" border="0" style="visibility:hidden">
<img id="BTNON_state" src="./images/BTN_ON.png" name="BTNON_state" width="5" height="5" border="0" style="visibility:hidden">
<img id="BTN_state" src="./images/off-2.png" name="BTN_state" width="5" height="5" border="0" style="visibility:hidden">


<a onMouseDown="" onMouseOver="" onMouseOut="">
	<img id="Light" src="images/light_OFF.png" name="Light"  border="1" width="125" height="117"  alt="Light" 
	style="border: 1px ridge #0066FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
</a>

<p>

<font size="6" color="#0000FF">current:
<span id="Current"> </span>
</font>
</p>
<img id="ImageOFF" src="./images/light_OFF.png" name="ImageOFF" width="5" height="5" border="0" style="visibility:hidden">
<img id="ImageON" src="./images/light_ON.png" name="ImageON" width="5" height="5" border="0" style="visibility:hidden">


<map name="Map">
  <area href="funsel.asp" shape="rect" coords="475, 20, 521, 44">
  <area href="floor1VW.asp" shape="rect" coords="546, 20, 589, 44">
  <area href="floor1AA.asp" shape="rect" coords="610, 16, 690, 44">
</map>
					
</BODY>
</HTML>
<Script Language="JavaScript">

    init();

</Script>
<%
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
if err.number <> 0 then
    response.write "["&sql&"]<br>"
    response.end
end if    
set open_recordset = rs
end function

function GetMdbRecordset(dbn, SQL)
dim cn,rs

set cn=db_connection(dbn)
set rs=open_recordset(cn,SQL,2,3)
set GetMdbRecordset=rs
end function
%>