#include "pch.h"
#include "jstring.h"
#include <stdio.h> 
#include <string>
#include "dbfile.h"
#include "readconfig.h"
int makeFieldData(char *s, char * s1, char *s2);
int makeMBFieldData(char *s, char * s1, char *s2, char **p, int *mbflag);
int strAG(char *s2, int n);
int StoI(char * s);
extern char **devNameOut;
extern char **devName;
extern int devN;
extern char *CnState;
extern Inputs *InP;
extern int recvFlag;
int clear_TAB(char *msg) 
{
	int i, j;
	for (i = 0,j=0;msg[i];i++) {
		if (msg[i] != 0x09) {
			msg[j++]=msg[i];
		}
	}
	msg[j] = 0x00;
	return j;
}
int CstrAg(char *s1, int n, char *s2);
int jstodb(char *topic, char * msg)
{
	int i,j,ii,t,id;
	char *p, *pf, s1[50], s2[550],dbn[50],*mbp;
	char buf[1000],sout[500],st[20];
	int k,rVal,MBFlag=0;
	string ts;
	
	p = NULL;recvFlag = 0;
	for (i = 0;i < devN;i++) {
		sprintf_s(s1, "/%s/", devName[i]);
		if ((p=strstr(topic, s1)) != NULL) {
			id = i;
			CnState[i] = 0;
			break;
		}
	}
	if (p == NULL) return 0;
	
	clear_TAB(msg);
//	printf("%s\n", msg);getchar();
	p = instr(msg, (char *)"{");
	if (p == NULL)return 0;
	
	i = 0;
	ts = "";dbn[0] = 0x00;
	MBFlag = 0;mbp = NULL;
	while (p = sgets(p, buf)) {
		s1[0] = 0x00;s2[0] = 0x00;
		if (strstr(buf, "modbus response") != NULL) {
			MBFlag = 1;
			mbp = buf;
		}
		if (MBFlag==1) {
			rVal = makeMBFieldData(mbp, s1, s2, &mbp,&MBFlag);
		}else
			rVal=makeFieldData(buf, s1, s2);

		if (rVal == 1 ) {
			
			if(strcmp(s1,"time")==0){
				if (s2[0] == '0'&&s2[1] == '0') {
					cmdInterpreter(devNameOut[id], (char *)"SynT");
					return 0;
				}
				ts = s1;
				ts += ":";
				ts += s2;
				ts += "|";

				j = 0;
				for (ii = 0;ii < 10;ii++) {
					if (s2[ii] != '/') {
						dbn[j] = s2[ii];
						j++;
					}
				}
				dbn[j] = 0x00;
			}

			if (strcmp(s1, "current data") == 0) {
				if (writefield(buf, s1, s2)) {
					ts += s1;
					ts += "_s:";
					ts += s2;
					ts += "|";
					CstrAg(s2, 15, st);
					ts += s1;
					ts += ":";
					ts += st;
					ts += "|";
				}

			}	

			for (ii = 0;ii < InP[id].length;ii++) {
				if (strcmp(InP[id].Fields[ii].Sname, s1) == 0) {	
//					printf("[%s][%s]\n", InP[id].Fields[ii].dbName,s2);
					pf = InP[id].Fields[ii].dbName;

					if ((pf = strstr(pf, ":")) != NULL) {
						switch (*(pf - 1)) {
						case 'S':
							pf++;
							updateDevState(devName[id], pf, s2);
							break;
						case 'A':
							pf++;
							CstrAg(s2, 15, st);
							ts += pf;
							ts += ":";
							ts += st;
							ts += "|";
							break;
						}
					}
					else {
						ts += InP[id].Fields[ii].dbName;
						ts += ":";
						ts += s2;
						ts += "|";
					}
				}
			}
		}
		if (p[0] == 0x00)break;		
	}

//	putsXY(sout, 60, 3);
	if (dbn[0] != 0x00) {
		ts += "<end>";
		write_recs_to_db(topic, dbn, ts);
	}

	recvFlag = 1;
//	puts(msg);
	return i;
}
int makeMBFieldData(char *s, char * s1, char *s2, char **p, int *mbflag)
{
	const char * keys[] = { "slave","function","start","length","out","error" };
	char *p1,*pe;
	int i,k=0;
	p1 = s;k = 0;
//	printf("[%s]", s);
	for (i = 0;i < 4;i++) {
		p1 = strstr(p1, keys[i]);
		if (p1 == NULL)return 0;
		if((p1= strstr(p1, ":"))==NULL) return 0;
		p1++;
		while (*p1 != ','&& *p1 !=0x0A && *p1)
			s1[k++] = *p1++;
		if(i!=3)s1[k++] = '#';
	}
	s1[k] = 0x00;

	if ((pe = strstr(p1, keys[5]) )==NULL) {
		p1 = strstr(p1, keys[4]);
		p1 = strstr(p1, "[");
		p1++;k = 0;
		while (*p1 != ']'&& *p1 != 0x0A && *p1)
			s2[k++] = *p1++;
		s2[k] = 0x00;
		*p = p1;
		if (strstr(*p, keys[0]) != NULL) {
			*mbflag = 1;
		}
		else {
			*mbflag = 0;
		}
	}else{
		s2[0] = 0x00;
		pe += strlen(keys[5]);
		*p = pe;
	}
	return 1;
}
int makeFieldData(char *s, char * s1, char *s2)
{
	int i;

	if (s[0] != '"')return 0;
	str_split(s, s1, s2);
	for (i = 0;s1[i];i++)
		if (s1[i] == ':') {
			s1[i] = 0x00;
			break;
		}
//	printf("[%s]:[%s]\n", s1, s2);
	return 1;
}
int writefield(char *s, char * s1,char *s2) 
{
	int i,t;

	if (s[0] != '"')return 0;
	t=str_split(s, s1, s2);
	for(i=0;s1[i];i++)
		if (s1[i] == ' ') {
			s1[i] = 0x00;
			break;
		}
//	printf("[%s]:[%s]", s1, s2);
	return t;
}
int writeAGfield(char *s,int n, char * s1, char *s2)
{
	int k=0,i,t;
	double v;

	if (s[0] != '"')return 0;
	t=str_split(s, s1, s2);
	for (i = 0;s1[i];i++)
		if (s1[i] == ' ') {
			s1[i] = 0x00;
			break;
		}
	CstrAg(s2, n, s2);

	return t;
}
int CstrAg(char *s1, int n, char *s2)
{
	int k;
	double v;

	k = strAG(s1, 15);
	v = k * 3.26;
	sprintf_s(s2, 50, "%8.1f", v);

	return 1;
}
int strAG(char *s, int n)
{
	int i,k,sum,I[100];
	char *ss[100];

	k = 0;
	ss[k] = s;k++;
	for (i = 0;s[i];i++) {
		if (s[i] == ',') {
			s[i] = 0x00;

			if (s[i + 1] != 0x00) {
				ss[k] = &(s[i + 1]);
				I[k-1] = i;
				k++;
			}
		}
	}
	sum = 0;
	for (i = k - n;i < k;i++) {
		sum += StoI(ss[i]);
	}
	for (i = 0;i < (k-1);i++) {
		s[I[i]]=',';
	}
	return sum;
}
int StoI(char * s)
{
	int i;
	int t = 0;

	for (i = 0;s[i];i++) {
		t = t * 10 + (s[i] - '0');
	}
	return t;
}
char * sgets(char * s, char * sout)
{
	int i;
	char *p;

	for (i = 0;s[i] && s[i] != 0x0A;i++)
		sout[i] = s[i];
	sout[i] = 0x00;
	if (strstr(sout, "modbus response")!=NULL) {
		do {
			for (;s[i] && s[i] != '}';i++) {
				sout[i] = s[i];
			}
			sout[i++] = '}';
		} while (s[i] == ',');

		for (;s[i] && s[i] != ']';i++) {
			sout[i] = s[i];
		}
	}

	sout[i] = 0x00;
	if (s[i] == 0x00)p = &(s[i]);
	else p = &(s[i + 1]);
	return p;
}
char *instr(char *s, char * s1)
{
	char * p;
	int n;
	char c;

	p = NULL;
	n = strlen(s1);
	//	printf("[%s][%d]\n",s1,n);getchar();
	while (strlen(s) >= n) {
		c = s[n];
		s[n] = 0x00;
		if (strcmp(s, s1) == 0) {
			p = s;
			s[n] = c;
			return p;
		}
		else {
			s[n] = c;
			s++;
		}
	}
	return p;
}

int str_split(char * buf, char *s1, char * s2)
{
	int i, j = 0,t;
	for (i = 0, j = 0;buf[i] != ':';i++) {
		if (buf[i] != '\"') {
			s1[j++] = buf[i];
		}
	}
	s1[j] = 0x00;t = j;
	i++;j = 0;
	for (;buf[i] != '\n' && buf[i];i++) {
		if (buf[i] != '\"'&&buf[i] != 'Z'&&buf[i] != '}'&&buf[i] != '\''&&buf[i] != '['&&buf[i] != ']') {
			s2[j++] = buf[i];
		}
	}
	if (s2[j - 1] == ',')j--;
	s2[j] = 0x00;

	return t;
}
