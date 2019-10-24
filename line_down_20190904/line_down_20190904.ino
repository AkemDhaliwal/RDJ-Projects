#include <SPI.h>
#include <DMD.h>
#include <TimedAction.h>
#include <TimerOne.h>
#include "SystemFont5x7.h"
#include "Arial_black_16.h"<arial_black_16.h>
#define DISPLAYS_ACROSS 1
#define DISPLAYS_DOWN 1
DMD dmd(DISPLAYS_ACROSS, DISPLAYS_DOWN);

String minsTens, minsUnits, minsThousand, minsHundred, value;
int mins10, mins1, mins100, mins1000, sec;
int minutes, seconds;
int clockRunning = 0;

int sensorpin4 = 4;                  //Switch between accumulative and one time mode
int sensorpin3 = 3;                 // digital pin used to connect the Roller sensor
int sensorpin2 = 5;                 // reset counter
int sensorpin5Roller = 2;
int sensorpin5RollerVal = 0;
int sensorpin3Val  = 0;                 // variable to store the values from sensor(initially zero)
int modeSwitch = 0;           // variable to read mode switch
int resetSwitch = 0;          //variable to read reset switch
int totalTime = 0;
bool dmdDisplayFLAG = false;
int Tmins1 = 0;
int Tmins10 = 0;
int Tmins100 = 0;
int Tmins1000 = 0;
bool TotalFlag = false;
int tempCount = 0;
void clockUpdate() {
  seconds++;
  if (seconds > 59) { 
    seconds = 0;
    minutes++;
  }
  minsThousand = minutes / 100;
  minsHundred = minutes / 10;
  mins10 = minutes % 10;
  
  mins1 = minutes;
  displayUpdate();
//call button inttrupt and sensor read
}

TimedAction timer = TimedAction(1000,clockUpdate);

void setup() {
  
  timer.disable();
  minutes = 0; seconds = 0;
  TotalFlag = false; 
  Serial.begin(9600);               // starts the serial monitor
  
  Timer1.initialize( 4000 );           //period in microseconds to call ScanDMD. Anything longer than 5000 (5ms) and you can see flicker.
  Timer1.attachInterrupt( ScanDMD );   //attach the Timer1 interrupt to ScanDMD which goes to dmd.scanDisplayBySPI()
  dmd.clearScreen( true );   //true is normal (all pixels off), false is negative (all pixels on)

  pinMode(sensorpin3,INPUT_PULLUP);
  pinMode(sensorpin2,INPUT_PULLUP);
  pinMode(sensorpin4,INPUT_PULLUP);
  pinMode(sensorpin5Roller,INPUT_PULLUP);
  dmd.selectFont( Arial_Black_16 );
}
void loop() {
  if (clockRunning == 0) {
      
      delay(1000); //see what effect this delay has on printing to display and time tracking
      //Serial.print(switchCount);
      sensorpin3Val   = digitalRead(sensorpin3);
      modeSwitch = digitalRead(sensorpin2);
      resetSwitch = digitalRead(sensorpin4);
     sensorpin5RollerVal = digitalRead(sensorpin5Roller); 
      if (resetSwitch == LOW)   
      {
          sec = 0;
          mins1000 = 0;
          mins100 = 0;
          mins100 = 0;
          mins10 = 0;
          mins1 = 0;
          Tmins1000 = 0;
          Tmins100 = 0;
          Tmins100 = 0;
          Tmins10 = 0;
          Tmins1 = 1;
      }
      if(modeSwitch == LOW)
      {
         if(TotalFlag == true)
          {
            TotalFlag = false;
            if (sensorpin3Val  == LOW)//check switch (sensor is up)
              {
                dmdDisplayFLAG = false;
              }
          }
          else
          {
            TotalFlag = true; 
            dmdDisplayFLAG = true; 
          }
      }
      //change this for new output
      if (sensorpin3Val  == HIGH)//check switch (sensor is up)
      {
        if(dmdDisplayFLAG == false)
        {
          mins1 = 1;
          dmdDisplayFLAG = true;
        }
        //mins1 = 1;
        tempCount++;
      }
      else
      {
        tempCount = 0;
        mins100 = 0;
        mins100 = 0;
        mins10 = 0;
        mins1 = 0;
        if(TotalFlag == false)
        {
          dmdDisplayFLAG = false;  
        }
      }
      if(tempCount > 0)
         {
            sec++;
            if(sec >= 59)
            {
              mins1++;
              Tmins1++;
              sec = 0;   
            }
            if(mins1 >= 10)
            {
              mins10++;
              Tmins10++;
              mins1 = 0;   
              Tmins1 = 0;
            }     
            if(mins10 >= 10) 
            {
              mins100++;
              Tmins100++;
              mins10 = 0;
              Tmins10 = 0;
            }
            if(mins100 >= 10)
            {
              mins1000++;
              Tmins1000++;
              mins100 = 0;
              Tmins100 = 0;
            }
            minutes = mins1000 * 1000 + mins100 * 100 + mins10 * 10 + mins1;
         }
         totalTime = totalTime + minutes;
         displayUpdate();
  }  
  if (clockRunning == 1) timer.check();  
}

void ScanDMD() { 
  dmd.scanDisplayBySPI();
}
void displayUpdate() {
  if(dmdDisplayFLAG == true)
  {   
    if (TotalFlag == false)
    {
      minsThousand = (String)mins1000;
      minsHundred = (String)mins100;
      minsTens = (String)mins10;
      minsUnits = (String)mins1;
    } 
    else 
    {
      minsThousand = (String)Tmins1000;
      minsHundred = (String)Tmins100;
      minsTens = (String)Tmins10;
      minsUnits = (String)Tmins1;
    }
  }
  else
  {
      minsThousand = " ";
      minsHundred = " ";
      minsTens =  " ";
      minsUnits = " ";
  }
  
  if(minsThousand == "0")
      minsThousand = " ";
  if(minsHundred == "0")
      minsHundred = " ";
  if(minsTens == "0")    
      minsTens =  " ";
      
  if(mins1 == 1 or mins1 == 4 or Tmins1 == 1 or Tmins1 == 4)
  {
      dmd.drawString(0,1,minsThousand.c_str(),1,GRAPHICS_NORMAL);
      dmd.drawString(8,1,minsHundred.c_str(),1,GRAPHICS_NORMAL);
      dmd.drawString(16,1,minsTens.c_str(),1,GRAPHICS_NORMAL);
      dmd.drawString(25,1,minsUnits.c_str(),1,GRAPHICS_NORMAL);
  }
  else
  {
    dmd.drawString(0,1,minsThousand.c_str(),1,GRAPHICS_NORMAL);
    dmd.drawString(8,1,minsHundred.c_str(),1,GRAPHICS_NORMAL);
    dmd.drawString(16,1,minsTens.c_str(),1,GRAPHICS_NORMAL);
    dmd.drawString(24,1,minsUnits.c_str(),1,GRAPHICS_NORMAL);
  }
}
