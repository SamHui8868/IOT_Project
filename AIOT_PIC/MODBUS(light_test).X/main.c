/**
  Generated Main Source File

  Company:
    Microchip Technology Inc.

  File Name:
    main.c

  Summary:
    This is the main file generated using PIC10 / PIC12 / PIC16 / PIC18 MCUs

  Description:
    This header file provides implementations for driver APIs for all modules selected in the GUI.
    Generation Information :
        Product Revision  :  PIC10 / PIC12 / PIC16 / PIC18 MCUs - 1.80.0
        Device            :  PIC18F23K22
        Driver Version    :  2.00
*/

/*
    (c) 2018 Microchip Technology Inc. and its subsidiaries. 
    
    Subject to your compliance with these terms, you may use Microchip software and any 
    derivatives exclusively with Microchip products. It is your responsibility to comply with third party 
    license terms applicable to your use of third party software (including open source software) that 
    may accompany Microchip software.
    
    THIS SOFTWARE IS SUPPLIED BY MICROCHIP "AS IS". NO WARRANTIES, WHETHER 
    EXPRESS, IMPLIED OR STATUTORY, APPLY TO THIS SOFTWARE, INCLUDING ANY 
    IMPLIED WARRANTIES OF NON-INFRINGEMENT, MERCHANTABILITY, AND FITNESS 
    FOR A PARTICULAR PURPOSE.
    
    IN NO EVENT WILL MICROCHIP BE LIABLE FOR ANY INDIRECT, SPECIAL, PUNITIVE, 
    INCIDENTAL OR CONSEQUENTIAL LOSS, DAMAGE, COST OR EXPENSE OF ANY KIND 
    WHATSOEVER RELATED TO THE SOFTWARE, HOWEVER CAUSED, EVEN IF MICROCHIP 
    HAS BEEN ADVISED OF THE POSSIBILITY OR THE DAMAGES ARE FORESEEABLE. TO 
    THE FULLEST EXTENT ALLOWED BY LAW, MICROCHIP'S TOTAL LIABILITY ON ALL 
    CLAIMS IN ANY WAY RELATED TO THIS SOFTWARE WILL NOT EXCEED THE AMOUNT 
    OF FEES, IF ANY, THAT YOU HAVE PAID DIRECTLY TO MICROCHIP FOR THIS 
    SOFTWARE.
*/

#include "mcc_generated_files/mcc.h"
#include "mcc_generated_files/eusart1.h"
#include"CD4021.h"
#include"CD4094.h"
#include <string.h>
/*
                         Main application
 */
uint16_t HR[8];
char MB_CMD_Buffer[32];     //receive buf
char MB_ACK_buf[16];        //send ack buf
uint8_t MBID;               //Local Address
uint8_t _4094_buf_H;
uint8_t _4094_buf_L;
extern uint32_t timecount;
int flagA=0;
int CountA = 0;

void Modbus_Read(void);
void HR_handler(void);              
void MB_handler(void);                     
void HR_read(void);                     //com03
void HR_single_write(void);              //cmd06
void write_ACK(uint8_t size); 
void Initial_Variable(void);
void DRIVER(void);
typedef union
{
    unsigned short v;
    struct
    {
        unsigned char LB ;
        unsigned char HB ;
    };
}SHORT;


void main(void)
{
    // Initialize the device
    SYSTEM_Initialize();

    // If using interrupts in PIC18 High/Low Priority Mode you need to enable the Global High and Low Interrupts
    // If using interrupts in PIC Mid-Range Compatibility Mode you need to enable the Global and Peripheral Interrupts
    // Use the following macros to:

    // Enable the Global Interrupts
    INTERRUPT_GlobalInterruptEnable();

    // Disable the Global Interrupts
    //INTERRUPT_GlobalInterruptDisable();

    // Enable the Peripheral Interrupts
    INTERRUPT_PeripheralInterruptEnable();

    // Disable the Peripheral Interrupts
    //INTERRUPT_PeripheralInterruptDisable();
    uint8_t data;
    uint8_t temp;
    uint32_t oldtimecount;
    oldtimecount= timecount;
    
    Initial_Variable();
    while (1)
    {
        Modbus_Read();
      
        EUSART1_Write_task(); 
        
        DRIVER();

        //temp = '2';
    

        
        // Add your application code
         /*if(timecount-oldtimecount>=10)
        {
        LED_1_LAT^=1;
        oldtimecount= timecount;
        }*/
    }
}

void Modbus_Read(void)
{
    static uint32_t MB_D_time = 0;        //modbus old count
    static uint8_t index = 0;              
    if(!Rx_isEmpty()) //read data from queue
    {
        MB_CMD_Buffer[index++] = EUSART1_Read();
        MB_D_time = timecount;
    }
        
    if ((uint32_t)(timecount - MB_D_time) > 3)   //check if timeout
    {     
        MB_handler();                   
        HR_handler();
        MB_D_time = timecount;
        index=0;                        
    }
}
void MB_handler(void)
{
    if (MB_CMD_Buffer[0] != MBID)
        return;
     switch (MB_CMD_Buffer[1])
    {
            case 03 : 
                HR_read();
                break;
            case 06 :
                HR_single_write();
                break;
    }            
}
void HR_read(void)
{   
    SHORT u ;
    uint8_t i;
    uint8_t j;
    uint16_t*p;
    uint16_t n;     //HR Read Count
    
    MB_ACK_buf[0] = MB_CMD_Buffer[0];
    MB_ACK_buf[1] = MB_CMD_Buffer[1];
    
    u.HB = MB_CMD_Buffer[2];
    u.LB = MB_CMD_Buffer[3];
    p = &(HR[u.v]);
    
    u.HB = MB_CMD_Buffer[4];
    u.LB = MB_CMD_Buffer[5];
    MB_ACK_buf[2] = u.v << 1;
    n= u.v ;
    
    j = 3;
    for (i=0;i<n;i++)
    {   
        u.v =p[i];
        MB_ACK_buf[j++] = u.HB;
        MB_ACK_buf[j++] = u.LB;
    }
    write_ACK(MB_ACK_buf[2]+3);
    
    memset(MB_CMD_Buffer,0x00,32);
    
}
void write_ACK(uint8_t size)
{
    uint8_t i ;
    for (i=0;i<size;i++)
    {
     EUSART1_Write(MB_ACK_buf[i]);
    }
}
void HR_single_write(void)
{   
    SHORT u;
    uint8_t i ;
    uint16_t addr;
    
    u.HB = MB_CMD_Buffer[2];
    u.LB = MB_CMD_Buffer[3];
    addr = u.v;
    u.HB = MB_CMD_Buffer[4];
    u.LB = MB_CMD_Buffer[5];
    HR[addr] = u.v;
    for(i=0;i<6;i++)
    {
      MB_ACK_buf[i] = MB_CMD_Buffer[i];
    }
    write_ACK(6);
    
    memset(MB_CMD_Buffer,0x00,32);
}

void HR_handler(void)
{
        int i;
    uint8_t data;
 for (i=0; i<8; i++) {
    data|= HR[i] << (i % 8);
                 }
    _4094_buf_H = 0xff;
    _4094_buf_L = data&0xff;   
}

void DRIVER(void)
{
        uint32_t CD4094DriveTimeCounter = 0;
         uint32_t CD4021DriveTimeCounter = 0;
         uint8_t data;
          bool _bit[8];
          data= CD4021Read();
          int i;
        for (i=0; i<8; i++) {
    _bit[i] = ((1 << (i % 8)) & data) >> (i % 8);
                            }
          if(_bit[0]==1)
          {
              CountA=0;
              flagA=1;
          }
          else
          {
              if(CountA>10)
              {
                  flagA=0;
              }
              CountA++;
          }
          if(flagA==1)
            HR[0]=1;
          else
             HR[0]=0;
          //HR[1]=_bit[1];
         // HR[2]=_bit[2];
          
         CD4094WriteH(_4094_buf_H);
         CD4094WriteL(_4094_buf_L);
     if((timecount - CD4094DriveTimeCounter)>=500){
            CD4094DriveTimeCounter = timecount;
            CD4094Driver();
        }
     if((timecount - CD4021DriveTimeCounter)>=500){
            CD4021DriveTimeCounter = timecount;
            CD4021Driver();
            LED_LAT^=1;
        }
}


void Initial_Variable(void)
{
    MBID = 0x03;
    
    for(int i=0;i < 8;i++)
    {
        HR[i] = 0;
    }
    
    _4094_buf_H =0xff;
    _4094_buf_L=0xff;
}



/**
 End of File
*/