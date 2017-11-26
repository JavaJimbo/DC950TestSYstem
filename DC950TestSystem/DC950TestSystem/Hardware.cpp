/* Hardware.cpp - mid level routines for operating Interface board, multimeter, and power supply.
 * All three devices use serial ports for communication, so SerialCom file is required
 * fpr low level routines.
 * Written in Visual C++ for DC950 test system
 *
 * 
 */

#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"
#include <iostream>
#include <sstream>
#include <windows.h>
#include "avaspec.h"
#include "BasicExcel.hpp"
#include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	

HANDLE gDoneEvent;
HANDLE hTimer = NULL;
HANDLE hTimerQueue = NULL;

extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

	void TestApp::closeAllSerialPorts(){
		closeSerialPort(handleInterfaceBoard); 
		closeSerialPort(handleHPmultiMeter);
		closeSerialPort(handleACpowerSupply);
	}

	float TestApp::getAbs(float floatValue) {
		float absoluteValue;
		if (floatValue < (float) 0.0) absoluteValue = (float) 0.0 - floatValue;
		else absoluteValue = floatValue;
		return absoluteValue;
	}

	VOID CALLBACK TimerRoutine(PVOID lpParam, BOOLEAN TimerOrWaitFired)
	{
		SetEvent(gDoneEvent);
	}

	void TestApp::msDelay(int milliseconds) {		
		gDoneEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
		hTimerQueue = CreateTimerQueue();
		int arg = 123;
		CreateTimerQueueTimer(&hTimer, hTimerQueue, (WAITORTIMERCALLBACK)TimerRoutine, &arg, milliseconds, 0, 0);
		WaitForSingleObject(gDoneEvent, INFINITE);
		CloseHandle(gDoneEvent);
	}
				
	BOOL TestApp::ReadVoltage(int multiplexerSelect, float *ptrMeasuredVoltage)
	{
		char strResponse[BUFFERSIZE] = "";
		char strReadVoltage[BUFFERSIZE] = ":MEAS?\r\n";		
		float floatVoltage;

		// Set relays on interface board for desired voltage source:
		if (!SetInterfaceBoardMulitplexer(multiplexerSelect)) return FALSE;

		msDelay(500);

		// Send READ VOLTAGE command to multimeter and get response:
		if (!sendReceiveSerial(HP_METER, strReadVoltage, strResponse, TRUE)) {
			DisplayMessageBox("HP METER COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		// If communication with multimeter was successful, convert response string to float value:
		else {
			floatVoltage = (float)atof(strResponse);
			*ptrMeasuredVoltage = getAbs(floatVoltage);
		}

		return TRUE;
	}
		
	BOOL TestApp::SetInterfaceBoardMulitplexer(int multiplexerSelect)
	{
		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE];

		switch (multiplexerSelect) {
		default:
		case LAMP:
			strcpy_s(strCommand, BUFFERSIZE, "$LAMP");
			break;
		case VREF:
			strcpy_s(strCommand, BUFFERSIZE, "$VREF");
			break;
		case CONTROL_VOLTAGE:		
			strcpy_s(strCommand, BUFFERSIZE, "$CTRL");
			break;
		}

		// Send RESET command to interface board:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("INTERFACE BOARD COMMUNICATION ERROR", "Check interface board cables and power", 1);
			return FALSE;
		}
		return TRUE;
	}

	BOOL TestApp::SetInterfaceBoardPWM(int PWMvalue)
	{
		char strCommand[BUFFERSIZE] = "$PWM>";
		char strValue[BUFFERSIZE], strResponse[BUFFERSIZE];

		sprintf_s(strValue, BUFFERSIZE, "%d", PWMvalue);
		strcat_s(strCommand, BUFFERSIZE, strValue);

		// Send SET PWM command to interface board:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("INTERFACE BOARD COMMUNICATION ERROR", "Check interface board cables and power", 1);
			return FALSE;
		}
		return TRUE;
	}
	

	BOOL TestApp::SetInhibitRelay(BOOL flagON)
	{
		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE];

		if (flagON) strcpy_s(strCommand, BUFFERSIZE, "$INHIBIT>ON");
		else strcpy_s(strCommand, BUFFERSIZE, "$INHIBIT>OFF");

		// Send RESET command to interface board:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("INTERFACE BOARD COMMUNICATION ERROR", "Check interface board cables and power", 1);
			return FALSE;
		}
		return TRUE;
	}
	


	BOOL TestApp::SetPowerSupplyVoltage(int commandVoltage) {
		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE] = "";
		char strReadVoltage[BUFFERSIZE] = "MEAS:VOLT?\r\n";		

		#ifdef SIMULMODE 
			return TRUE; 
		#endif

    	  // Send SET VOLTAGE command:
		sprintf_s(strCommand, "VOLT %d\r\n", commandVoltage);

		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
			
		// Otherwise, return TRUE to indicate power supply is communicating and functioning:
		return TRUE;
	}

	BOOL TestApp::SetInterfaceBoardActuatorOutput(BOOL actuatorClosed)
	{
		char strCommand[BUFFERSIZE] = "$TTL_HIGH";
		char strInputCommand[BUFFERSIZE] = "$TTL_IN";
		char strResponse[BUFFERSIZE];

		SetInterfaceBoardPWM(MAX_PWM);
		msDelay(100);

		if (actuatorClosed == TRUE) strcpy_s(strCommand, BUFFERSIZE, "$TTL_LOW");
		else strcpy_s(strCommand, BUFFERSIZE, "$TTL_HIGH"); 

		// Send SET PWM command to interface board:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("INTERFACE BOARD COMMUNICATION ERROR", "Check interface board cables and power", 1);
			return FALSE;
		}

		// Check TTL input for actuator fault detect:
		if (!sendReceiveSerial(INTERFACE_BOARD, strInputCommand, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("DC950 TTL INPUT ERROR", "The DC950 has failed the TTL input test", 1);
			return FALSE;
		}

		return TRUE; // SUCCESS! Both the interface board and the DC950 are responding to commands
	}
	
