/* Initialize.cpp - routines for initializing hardware at startup and also reading/writing INI file.
 * Includes functions for initializing interface board, multimeter, and power supply.
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

extern float AllowableLampVoltageError;
extern float AllowableVrefError;
extern float MinClosedFilterColorTemp;
extern float MaxClosedFilterColorTemp;
extern float MinClosedFilterWavelength;
extern float MaxClosedFilterWavelength;
extern float MinClosedFilterIrradiance;
extern float MaxClosedFilterIrradiance;
extern float MinClosedFilterFWHM;
extern float MaxClosedFilterFWHM;
extern float MinClosedFilterAmplitude;

extern float MinOpenColorTemp;
extern float MaxOpenColorTemp;
extern float MinOpenIrradiance;
extern float MaxOpenIrradiance;

// #define MAXINIVALUES 32

extern const char *INIfilename;
extern float arrINIconfigValues[];
extern long portNumberInterfaceBoard, portNumberACpowerSupply, portNumberMultiMeter;
extern float wavelength, minAmplitude, peakIrradiance, FWHM, colorTemperature;
extern int currentRow;
extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

	void TestApp::resetTestData() {
		for (int i = 0; i < FINAL_PASS; i++) 
			testStep[i].Status = NOT_DONE_YET;		
	}
	
	// This is called once, when ENTER button
	// has been pushed the first time.
	UINT TestApp::SystemStartUp(CTestDialog *ptrDialog)  
	{
		if (!InitializeSystem(ptrDialog)) return (FAIL);		
		else {			
			resetDisplays(ptrDialog);			
			return (PASS);
		}
	}


	BOOL TestApp::InitializeHP34401(CTestDialog *ptrDialog) {
		char strReset[BUFFERSIZE] = "*RST\r\n";
		char strEnableRemote[BUFFERSIZE] = ":SYST:REM\r\n";
		char strMeasure[BUFFERSIZE] = ":MEAS?\r\n";

		// 1) Send RESET command to HP34401:
		if (!sendReceiveSerial(HP_METER, ptrDialog, strReset, NULL, FALSE)) {
			DisplayMessageBox("HP MULTIMETER COM ERROR", "Make sure meter power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}

		msDelay(500);

		// 2) Enable RS232 remote control on HP34401:
		if (!sendReceiveSerial(HP_METER, ptrDialog, strEnableRemote, NULL, FALSE)) {
			DisplayMessageBox("HP MULTIMETER COM ERROR", "Make sure meter power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}

		msDelay(500);

		// Now try getting a measurement from the HP34401:
		if (!sendReceiveSerial(HP_METER, ptrDialog, strMeasure, NULL, TRUE)) {
			DisplayMessageBox("HP MULTIMETER COM ERROR", "Make sure meter power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}

		return TRUE;
	}

	BOOL TestApp::InitializeInterfaceBoard(CTestDialog *ptrDialog) {
		char strReset[BUFFERSIZE] = "$RESET";
		char strResponse[BUFFERSIZE] = "";

		// Send RESET command to interface board:
		if (!sendReceiveSerial(INTERFACE_BOARD, ptrDialog, strReset, strResponse, TRUE)) {
			DisplayMessageBox("INTERFACE BOARD COM ERROR", "Make sure Interface board power is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		else if (!strstr(strResponse, "OK")) {
			DisplayMessageBox("INTERFACE BOARD COMMUNICATION ERROR", "Check interface board cables and power", 1);
			return FALSE;
		}
		return TRUE;
	}

	BOOL TestApp::InitializePowerSupply(CTestDialog *ptrDialog) {
		char strCommand[BUFFERSIZE] = "*RST\r";

		// Send RESET commands to programmable power supply:
		strcpy_s(strCommand, BUFFERSIZE, "*RST\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}		
		msDelay(1000);
		
		strcpy_s(strCommand, BUFFERSIZE, "*CLS\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		msDelay(1000);

		strcpy_s(strCommand, BUFFERSIZE, "*SRE 128\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		msDelay(1000);

		strcpy_s(strCommand, BUFFERSIZE, "*ESE 0\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		msDelay(1000);		

		strcpy_s(strCommand, BUFFERSIZE, "VOLT:RANG 272\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
		
		msDelay(1000);

		strcpy_s(strCommand, BUFFERSIZE, "OUTP 1\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}
				
		msDelay(1000);

		strcpy_s(strCommand, BUFFERSIZE, "SYST:REM\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}		

		msDelay(1000);

		strcpy_s(strCommand, BUFFERSIZE, "VOLT 120\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, ptrDialog, strCommand, NULL, FALSE)) {
			DisplayMessageBox("AC POWER SUPPLY COM ERROR", "Make sure AC power supply is ON, \r\nCheck RS232 and USB cables", 1);
			return FALSE;
		}

		msDelay(1000);

		// Return TRUE to indicate power supply is communicating and functioning:		
		return TRUE;
	}
	
	
	BOOL TestApp::InitializeSystem(CTestDialog *ptrDialog) {
		char portNameInterfaceBoard[MAXSTRING], portNameACpowerSupply[MAXSTRING], portNameMultiMeter[MAXSTRING];
		
		portNumberInterfaceBoard = (long) arrINIconfigValues[0];
		if (portNumberInterfaceBoard < 1 || portNumberInterfaceBoard > MAX_PORTNUMBER)
			return FALSE;

		portNumberACpowerSupply = (long) arrINIconfigValues[1];
		if (portNumberACpowerSupply < 1 || portNumberACpowerSupply > MAX_PORTNUMBER)
			return FALSE;

		portNumberMultiMeter = (long) arrINIconfigValues[2];
		if (portNumberMultiMeter < 1 || portNumberMultiMeter > MAX_PORTNUMBER)
			return FALSE;
		
		sprintf_s(portNameInterfaceBoard, MAXSTRING, "\\\\.\\COM%d", (int) portNumberInterfaceBoard);
		sprintf_s(portNameACpowerSupply, MAXSTRING, "\\\\.\\COM%d", (int) portNumberACpowerSupply);
		sprintf_s(portNameMultiMeter, MAXSTRING, "\\\\.\\COM%d", (int) portNumberMultiMeter);
				
		if (openSerialPort(portNameInterfaceBoard, &handleInterfaceBoard)
			&& openSerialPort(portNameMultiMeter, &handleHPmultiMeter)
			&& openSerialPort(portNameACpowerSupply, &handleACpowerSupply)) {
			ptrDialog->m_static_DataOut.SetWindowText((LPCTSTR)"COM Data Out");
			ptrDialog->m_static_DataIn.SetWindowText((LPCTSTR)"COM Data In");			
			if (!InitializeHP34401(ptrDialog)) return FALSE;
			if (!InitializeInterfaceBoard(ptrDialog)) return FALSE;
			
#ifndef SIMULMODE
			if (!InitializePowerSupply(ptrDialog)) return FALSE;			
			if (!InitializeSpectrometer(ptrDialog)) return FALSE:
#endif			
			return TRUE;
		}
		else return FALSE;
	}

	
void TestApp::write_INI_File(){	
	float INIconfigValue = 0;
	int i = 0;

	arrINIconfigValues[0] = (float) portNumberInterfaceBoard;
	arrINIconfigValues[1] = (float) portNumberACpowerSupply;
	arrINIconfigValues[2] = (float) portNumberMultiMeter;

	arrINIconfigValues[3] = AllowableLampVoltageError;
	arrINIconfigValues[4] = AllowableVrefError;
	arrINIconfigValues[5] = MinClosedFilterColorTemp;
	arrINIconfigValues[6] = MaxClosedFilterColorTemp;
	arrINIconfigValues[7] = MinClosedFilterWavelength;
	arrINIconfigValues[8] = MaxClosedFilterWavelength;
	arrINIconfigValues[9] = MinClosedFilterIrradiance;
	arrINIconfigValues[10] = MaxClosedFilterIrradiance;
	arrINIconfigValues[11] = MinClosedFilterFWHM;
	arrINIconfigValues[12] = MaxClosedFilterFWHM;
	arrINIconfigValues[13] = MinClosedFilterAmplitude;
	arrINIconfigValues[14] = MinOpenColorTemp;
	arrINIconfigValues[15] = MaxOpenColorTemp;
	arrINIconfigValues[16] = MinOpenIrradiance;
	arrINIconfigValues[17] = MaxOpenIrradiance;

	std::ofstream outFile;
	// Open file to create it
	outFile.open(INIfilename, ios::out|ios::binary|ios::trunc);
	if (!outFile.is_open()) 
		throw ios::failure(string("Error opening INI file ") + 
	string(INIfilename) +
	string("  in main()"));
	i = 0;
	do {
		INIconfigValue = arrINIconfigValues[i];
		writeBinaryFile(outFile, INIconfigValue);
		i++;
	} while (i < MAXINIVALUES);
	outFile.close();				
}



BOOL TestApp::load_INI_File(){
	std::ifstream inFile;
	float INIconfigValue = 0;

	inFile.open(INIfilename, ios::in|ios::binary); // Open INI file		
	if (inFile.is_open()){
		int i = 0;
		do {
			readBinaryFile(inFile, INIconfigValue);
			arrINIconfigValues[i] = INIconfigValue;
			i++;
		} while (!inFile.eof() && i < MAXINIVALUES);
		portNumberInterfaceBoard = (long) arrINIconfigValues[0];
		portNumberACpowerSupply = (long) arrINIconfigValues[1];
		portNumberMultiMeter = (long) arrINIconfigValues[2];	
		AllowableLampVoltageError = arrINIconfigValues[3];
		AllowableVrefError = arrINIconfigValues[4];
		MinClosedFilterColorTemp = arrINIconfigValues[5];
		MaxClosedFilterColorTemp = arrINIconfigValues[6];
		MinClosedFilterWavelength = arrINIconfigValues[7];
		MaxClosedFilterWavelength = arrINIconfigValues[8];
		MinClosedFilterIrradiance = arrINIconfigValues[9];
		MaxClosedFilterIrradiance = arrINIconfigValues[10];
		MinClosedFilterFWHM = arrINIconfigValues[11];
		MaxClosedFilterFWHM = arrINIconfigValues[12];
		MinClosedFilterAmplitude = arrINIconfigValues[13];
		MinOpenColorTemp = arrINIconfigValues[14];
		MaxOpenColorTemp = arrINIconfigValues[15];
		MinOpenIrradiance = arrINIconfigValues[16];
		MaxOpenIrradiance = arrINIconfigValues[17];
		inFile.close();
	}
	else {
		write_INI_File();
		DisplayMessageBox("SYSTEM ERROR", "No INI file, creating default INI.", 1);	
		return TRUE;
	}
	return TRUE;
}

// Read binary float value
void TestApp::readBinaryFile(std::istream& in, float& value){
	in.read(reinterpret_cast<char *>(&value), sizeof(value));
}
// Write binary float value
void TestApp::writeBinaryFile(std::ostream& out, float value){
	out.write(reinterpret_cast<char *>(&value), sizeof(value));
}


