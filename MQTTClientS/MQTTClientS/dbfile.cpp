#include "pch.h"
#include <stdio.h> 
#include <iostream> 
#include <sstream>
#include <process.h>
#include <conio.h>
#include <fstream>
#include <string>
#include <time.h>
#include <direct.h>
#include <sys/stat.h>
using namespace std;

#import "c:\program files\common files\system\ado\msado15.dll" no_namespace rename ("EOF", "adoEOF")
#include <Windows.h>
void Fcopy(char * buf, int len, _variant_t vs);
int TestAndSet_db(void);
void ReleaseTAS_db(void);
int TestAndSet_db1(void);
void ReleaseTAS_db1(void);
int TestAndSet_db2(void);
void ReleaseTAS_db2(void);
extern int ExitServer;
int clear_slash(char *s,char *s1)
{
	int n = 0, i = 0;
	const char *cs = "/*";
	while (s[i]) {
		if (strchr(cs, s[i]) != NULL) {
			s1[i] = '_';
			n++;
		}
		else {
			s1[i] = s[i];
		}
		i++;
	}
	s1[i] = 0x00;
	return n;
}

int fileExist(char * pfln)
{
	struct _stat buf;
	int result;

	result = _stat(pfln, &buf);

	if (result != 0) return 0;

	return 1;
}
int connectAccdb(char *dbn, _ConnectionPtr *cnn)
{
	char cnstr[256];
	static HRESULT iniOK = 0;
	HRESULT hr;
	_ConnectionPtr m_pConnection;
	_variant_t      RecordsAffected;
	_RecordsetPtr   pRecordset = NULL;

	*cnn = 0x00;
	if (dbn[0] == 0x00)
		return 0;

	if (fileExist(dbn)) {
		iniOK = CoInitialize(NULL);

		if (iniOK == S_OK) {
			//			hr = m_pConnection.CreateInstance("ADODB.Connection");
			//			sprintf(cnstr,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= \"%s\" ;Jet OLEDB:DataBase Password=;",dbn);
			//			m_pConnection->Open(cnstr,"","",adModeUnknown); 

			try {
				hr = m_pConnection.CreateInstance("ADODB.Connection");///創建Connection對象

				if (SUCCEEDED(hr))
				{
					//					sprintf_s(cnstr,256,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= \"%s\" ",dbn);
					sprintf_s(cnstr, 256, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= \"%s\" ;Persist Security Info=False;", dbn);
					m_pConnection->Open(cnstr, "", "", adModeUnknown);
				}
			}
			catch (_com_error e)///捕捉異常 
			{
				printf("DB Connected Error![%s]", cnstr); getchar();
				CoUninitialize();
				return 0;
			}

			*cnn = m_pConnection;

			return 1;
		}
	}
	return 0;
}
int connectMdb(char *dbn, _ConnectionPtr *cnn);

int connectdb(char *dbn, _ConnectionPtr *cnn)
{

	if (strstr(dbn, "mdb") != NULL) {
		return connectMdb(dbn, cnn);
	}
	if (strstr(dbn, "accdb") != NULL) {
		return connectAccdb(dbn, cnn);
	}

	return 0;
}
int connectMdb(char *dbn, _ConnectionPtr *cnn)
{
	char cnstr[256];
	static HRESULT iniOK = 0;
	HRESULT hr;
	_ConnectionPtr m_pConnection;
	_variant_t      RecordsAffected;
	_RecordsetPtr   pRecordset = NULL;

	*cnn = 0x00;
	if (dbn[0] == 0x00)
		return 0;

	if (fileExist(dbn)) {
		iniOK = CoInitialize(NULL);

		if (iniOK == S_OK) {
			//			hr = m_pConnection.CreateInstance("ADODB.Connection");
			//			sprintf(cnstr,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= \"%s\" ;Jet OLEDB:DataBase Password=;",dbn);
			//			m_pConnection->Open(cnstr,"","",adModeUnknown); 

			try {
				hr = m_pConnection.CreateInstance("ADODB.Connection");///創建Connection對象

				if (SUCCEEDED(hr))
				{
					sprintf_s(cnstr, 256, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= \"%s\" ", dbn);
					//			sprintf(cnstr,"Microsoft.ACE.OLEDB.12.0;Data Source= \"%s\" ;Persist Security Info=False;",dbn);
					m_pConnection->Open(cnstr, "", "", adModeUnknown);
				}
			}
			catch (_com_error e)///捕捉異常 
			{
//				printf("DB Connected Error![%s]", cnstr); getchar();
				CoUninitialize();
				return 0;
			}

			*cnn = m_pConnection;

			return 1;
		}
	}
	return 0;
}
int openRecordset(_ConnectionPtr m_pConnection, char *SQL, _RecordsetPtr *rs)
{
	_variant_t      RecordsAffected;
	_RecordsetPtr   pRecordset = NULL;

	*rs = 0x00;
	if (m_pConnection != 0x00) {
		pRecordset.CreateInstance(_uuidof(Recordset));
		pRecordset->Open(SQL, _variant_t((IDispatch *)m_pConnection, true), adOpenStatic, adLockOptimistic, adCmdText);

		*rs = pRecordset;
		return 1;

	}
	return 0;
}
int create_db_path(char *dirn, char * dbn);
int write_recs_to_db(char *topic, char * dbn, string ts)
{
	int count = 0, state = 0,i,ii;
	char buf[1000], *s1, *s2, dbpath[256],fieldName[500],*fdn,stt[40];
	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	istringstream iss(ts);
	string s;
	
	while (TestAndSet_db()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	clear_slash(topic,buf);
	sprintf_s(dbpath, 256, "%s\\%s.mdb", buf, dbn);
	if (!fileExist(dbpath)) {
		create_db_path(buf,dbn);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);

	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db();
		return 0;
	}
//	printf("\n[%s]", dbpath);
	m_pConnection->BeginTrans();
	sprintf_s(buf, 100, "select * from DATA");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);
	fieldName[0]=0x00;
	for (i = 0;i < pRecordset->GetFields()->Count;i++) {
		s = pRecordset->GetFields()->GetItem(long(i))->Name;
		ii = strlen(fieldName);		
		sprintf_s((char *)&(fieldName[ii]),40, "[%s]", s.c_str());
	}

	while (getline(iss, s, '|')) {
		pRecordset->AddNew();
		do {
			
			strcpy_s(buf, 1000, s.c_str());

			if (strstr(buf, "<end>") != NULL)break;

			for (i = 0;buf[i] && buf[i] != ':';i++);
			buf[i] = 0x00;
			s1 = buf; s2 = &(buf[i + 1]);
//			printf("[%s]", s1);
			sprintf_s(stt,40, "[%s]", s1);
			if (strstr(fieldName, stt) != NULL) {
				pRecordset->PutCollect(s1, _variant_t((char *)s2));
			}

		} while (getline(iss, s, '|'));
		//		pRecordset->PutCollect("Feed_ID", _variant_t((char *)Feed_ID));
		//		pRecordset->PutCollect("Channel_ID", _variant_t((char *)Channel_ID));

		pRecordset->Update();
		count++;
	}

	pRecordset->Close();

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db();
	return count;
}
int read_recs_from_db(char *topic, char * dbn, char * fieldName, char *fieldData)
{
	int count = 0, state = 0, i;
	char buf[1000], *s1, *s2, dbpath[256];
	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	_variant_t vs,ii;
	string  s;

	fieldData[0] = 0x00;
	while (TestAndSet_db()) {
		if (ExitServer) return -1;
		Sleep(1);
	}
	clear_slash(topic, buf);
	sprintf_s(dbpath, 256, "%s\\%s.mdb", buf, dbn);

	if (!fileExist(dbpath)) {
		return 0;
	}
	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);

	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db();
		return 0;
	}
	//	printf("\n[%s]", dbpath);
	m_pConnection->BeginTrans();
	sprintf_s(buf, 100, "select * from DATA");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);
	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		pRecordset->MoveLast();
		for (i = 0;i < pRecordset->GetFields()->Count;i++) {
			s = pRecordset->GetFields()->GetItem(long(i))->Name;
			if (strcmp(s.c_str(), fieldName) == 0) {
				vs = pRecordset->GetFields()->GetItem(long(i))->GetValue();
				if (vs.vt != VT_NULL) {
					s = (LPCSTR)_bstr_t(vs);
					for (i = 0;(s.c_str())[i];i++)
						fieldData[i] = (s.c_str())[i];
					fieldData[i] = 0x00;

				}
				break;
			}
		}
		pRecordset->Close();
	}

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db();
	return state;
}
//====================================================================================
int delActiveSch(char *swName, int WeekDay, char * Stime, char * Atype)
{
	int state,count=0;
	char buf[200], dbpath[200];

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;

	sprintf_s(dbpath, 200, "devActive.mdb");

	while (TestAndSet_db1()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);

	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath,count);
		ReleaseTAS_db1();
		return 0;
	}
	m_pConnection->BeginTrans();
	sprintf_s(buf, 200, "delete from ActTable where Name=\"%s\" and Stime=\"%s\" and AType=\"%s\" and WeekDay=\"%d\"  ",
		      swName, Stime, Atype, WeekDay);

	m_pConnection->Execute(buf, NULL, adCmdText);

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db1();
	return state;
}

int getActiveSch(char *swName, int WeekDay, string * tsp)
{
	int count = 0, state;
	char buf[200], dbpath[200], s1[100], s2[100], s3[100], s4[100];

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	_variant_t vs;

	string ts="";

	sprintf_s(dbpath, 200, "devActive.mdb");
	while (TestAndSet_db1()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);

	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db1();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from ActTable where Name=\"%s\" and WeekDay=\"%d\" ", swName,WeekDay);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		while (!pRecordset->adoEOF) {
			Fcopy(s1, 100, pRecordset->GetCollect(_bstr_t("Stime")));
			Fcopy(s2, 100, pRecordset->GetCollect(_bstr_t("Etime")));
			Fcopy(s3, 100, pRecordset->GetCollect(_bstr_t("Aflag")));
			Fcopy(s4, 100, pRecordset->GetCollect(_bstr_t("Atype")));
			sprintf_s(buf,200,"%s/%s/%s/%s", s1, s2, s3,s4);
			if (ts == "") {
				ts = buf;
			}
			else {
				ts += "|";ts += buf;
			}
			pRecordset->MoveNext();
		}
		pRecordset->Close();
	}
	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();

	*tsp = ts;

	ReleaseTAS_db1();

	return 1;
}
void Fcopy(char * buf, int len, _variant_t vs)
{
	string  vsBuf;
	buf[0] = 0x00;
	if (vs.vt != VT_NULL) {
		vsBuf = (LPCSTR)_bstr_t(vs);
		strcpy_s(buf, len, vsBuf.c_str());
	}
}
//============================================================================
int getSWState(char *swName, char *cmd)
{
	int count = 0, state;
	char buf[300], dbpath[200];
	string  dname;
	int i;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	_variant_t vs;

	sprintf_s(dbpath,200,"dev.mdb");
	cmd[0] = '\0';
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);

	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db2();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 300, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		vs = pRecordset->GetCollect(_bstr_t("swState"));
		if(vs.vt != VT_NULL){
		    dname = (LPCSTR)_bstr_t(vs);
			for (i = 0;(dname.c_str())[i];i++)
				cmd[i] = (dname.c_str())[i];
			cmd[i] = 0x00;

		}
		pRecordset->Close();
	}

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}
int getSWcmd(char *swName, char *cmd)
{
	int count = 0, state;
	char buf[200], dbpath[200];
	string  dname;
	int i;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	_variant_t vs;

	sprintf_s(dbpath, 200, "dev.mdb");
	cmd[0] = '\0';
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);
	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db2();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		vs = pRecordset->GetCollect(_bstr_t("switchCTL"));
		if (vs.vt != VT_NULL) {
			dname = (LPCSTR)_bstr_t(vs);
			for (i = 0;(dname.c_str())[i];i++)
				cmd[i] = (dname.c_str())[i];
			cmd[i] = 0x00;

		}
		pRecordset->Close();
	}

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}
int writeSWcmd(char *swName, char *cmd)
{
	int count = 0, state;
	char buf[200], dbpath[200];
	string  dname;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;

	sprintf_s(dbpath, 200, "dev.mdb");
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);
	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db2();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		pRecordset->PutCollect("switchCTL", _variant_t((char *)cmd));
		pRecordset->Update();
	}

	pRecordset->Close();

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}
//=============================================================
int getClientcmd(char *swName, char *cmd)
{
	int count = 0, state;
	char buf[200], dbpath[200];
	string  dname;
	int i;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;
	_variant_t vs;

	sprintf_s(dbpath, 200, "dev.mdb");
	cmd[0] = '\0';
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);
	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db2();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		vs = pRecordset->GetCollect(_bstr_t("cmd"));
		if (vs.vt != VT_NULL) {
			dname = (LPCSTR)_bstr_t(vs);
			for (i = 0;(dname.c_str())[i];i++)
				cmd[i] = (dname.c_str())[i];
			cmd[i] = 0x00;

		}
		pRecordset->Close();
	}

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}
int writeClientcmd(char *swName, char *cmd)
{
	int count = 0, state;
	char buf[200], dbpath[200];
	string  dname;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;

	sprintf_s(dbpath, 200, "dev.mdb");
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);
	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		ReleaseTAS_db2();
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		pRecordset->PutCollect("cmd", _variant_t((char *)cmd));
		pRecordset->Update();
	}

	pRecordset->Close();

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}

//============================================================
int updateDevState(char *swName, char * atr, char *data)
{
	int count = 0, state;
	char buf[200], dbpath[200];
	string  dname;

	_ConnectionPtr  m_pConnection;
	_RecordsetPtr   pRecordset = NULL;
	_variant_t		RecordsAffected;

	sprintf_s(dbpath, 200, "dev.mdb");
	while (TestAndSet_db2()) {
		if (ExitServer) return -1;
		Sleep(1);
	}

	do {
		state = connectdb(dbpath, (_ConnectionPtr *)&m_pConnection);
		if (state == 0)Sleep(5);
		count++;
	} while (state == 0 && count <= 10);
	if (state == 0) {
		printf("DB Connected Error![%s][%d]\n", dbpath, count);
		return 0;
	}
	m_pConnection->BeginTrans();

	sprintf_s(buf, 200, "select * from Devices where Name=\"%s\"  ", swName);
	//    sprintf_s(buf,100,"select * from BasicData ");
	state = openRecordset(m_pConnection, buf, (_RecordsetPtr *)&pRecordset);

	if (!(pRecordset->adoEOF && pRecordset->BOF)) {
		pRecordset->PutCollect(atr, _variant_t((char *)data));
		pRecordset->Update();
	}

	pRecordset->Close();

	m_pConnection->CommitTrans();

	m_pConnection->Close();
	CoUninitialize();
	ReleaseTAS_db2();
	return 1;
}

int create_db_path(char *dirn, char * dbn)
{
	char cmd[300];

	_mkdir(dirn);
	sprintf_s(cmd, 300, "copy db_src.mdb %s\\%s.mdb\n", dirn,dbn);

	system(cmd);
	Sleep(500);
	return 1;
}

//------------------------------------------------------------------------------------------------------------------
static long TstFlag_db = 0;
int TestAndSet_db(void)
{
	long tmp;
	tmp = 1;
	_asm {
		MOV EAX, tmp
		xchg TstFlag_db, EAX
		MOV tmp, EAX
	}
	return tmp;
}
void ReleaseTAS_db(void)
{
	TstFlag_db = 0;
}
static long TstFlag_db1 = 0;
int TestAndSet_db1(void)
{
	long tmp;
	tmp = 1;
	_asm {
		MOV EAX, tmp
		xchg TstFlag_db1, EAX
		MOV tmp, EAX
	}
	return tmp;
}
void ReleaseTAS_db1(void)
{
	TstFlag_db1 = 0;
}
static long TstFlag_db2 = 0;
int TestAndSet_db2(void)
{
	long tmp;
	tmp = 1;
	_asm {
		MOV EAX, tmp
		xchg TstFlag_db2, EAX
		MOV tmp, EAX
	}
	return tmp;
}
void ReleaseTAS_db2(void)
{
	TstFlag_db2 = 0;
}
