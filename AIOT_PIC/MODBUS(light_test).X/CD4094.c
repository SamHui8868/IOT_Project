/*
 * File:   _4094.c
 * Author: user
 *
 * Created on March 19, 2019, 12:42 PM
 */


#include "CD4094.h"
#include "mcc_generated_files/mcc.h"
#include "mcc_generated_files/pin_manager.h"
union Buffer16{
    struct{
        uint8_t BufferH;
        uint8_t BufferL;
    };
    uint16_t Buffer;
};
static union Buffer16 CD4094Buffer;

void CD4094Driver(){
    //print data from CD4094Buffer.Buffer to 4094
    uint16_t temp;
    temp=CD4094Buffer.Buffer;
    CD4094_Data_LAT=0;
    NOP();
    NOP();
    CD4094_Strobe_LAT=0;
    NOP();
    NOP();
    CD4094_Clock_LAT=0;
    for(int i=0;i<16;i++){
        NOP();
        NOP();
        CD4094_Data_LAT=(temp&0x0001);
        CD4094_Clock_LAT=1;
        NOP();
        NOP();
        CD4094_Clock_LAT=0;
        if(i==15){
            CD4094_Strobe_LAT=1;
            NOP();
            NOP();
            CD4094_Strobe_LAT=0;
        }
        temp>>=1;
    }
    return;
}
void CD4094Write16(uint16_t Buffer){
    //put Buffer into CD4094Buffer.Buffer
    CD4094Buffer.Buffer=Buffer;
    //return Buffer;
}
void CD4094WriteH(uint8_t v){
    //put BufferH into CD4094Buffer.BufferH
    CD4094Buffer.BufferH=v;
    //return BufferH;
}
void CD4094WriteL(uint8_t s){
    //put BufferH into CD4094Buffer.BufferH
    

    CD4094Buffer.BufferL=s;
    
    //return BufferH;
}
