#pragma once
#include <sstream>
using namespace std;

int fileExist(char * pfln);
int write_recs_to_db(char *topic, char * dbn, string ts);
int read_recs_from_db(char *topic, char * dbn, char * fieldName, char *fieldData);
int writeSWcmd(char *swName, char *cmd);
int getSWcmd(char *swName, char *cmd);
int updateDevState(char *swName, char * atr, char *state);
int getSWState(char *swName, char *cmd);
int delActiveSch(char *swName, int WeekDay, char * Stime, char * Atype);
int getActiveSch(char *swName, int WeekDay, string * tsp);
int getClientcmd(char *swName, char *cmd);
int writeClientcmd(char *swName, char *cmd);