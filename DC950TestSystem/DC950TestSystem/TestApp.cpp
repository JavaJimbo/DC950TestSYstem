/* TestApp.cpp - for DC950 test system
 * 
 * This file includes Execute(), which implements the test sequence
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

float AllowableLampVoltageError = (float) 0.1;
float AllowableVrefError = (float) 0.1;
float MinClosedFilterColorTemp = (float) 100.0;
float MaxClosedFilterColorTemp = (float) 200.0;
float MinClosedFilterWavelength = (float) 938.0;
float MaxClosedFilterWavelength = (float) 943.0;
float MinClosedFilterIrradiance = (float) 200.0;
float MaxClosedFilterIrradiance = (float) 300.0;
float MinClosedFilterFWHM = (float) 10.0;
float MaxClosedFilterFWHM = (float) 18.0;
float MinClosedFilterAmplitude = (float) 5;

float MinOpenColorTemp = (float) 100.0;
float MaxOpenColorTemp = (float) 200.0;
float MinOpenIrradiance = (float) 200.0;
float MaxOpenIrradiance = (float) 300.0;

const char *INIfilename = "C:\\DC950_Data\\INIfile.bin";
float arrINIconfigValues[MAXINIVALUES];
long portNumberInterfaceBoard = 3, portNumberACpowerSupply = 4, portNumberMultiMeter = 7;


float wavelength = 940, minAmplitude = 5, peakIrradiance = 250, FWHM = 14, colorTemperature = 150;
int currentRow = 0;
HANDLE handleInterfaceBoard = NULL, handleHPmultiMeter = NULL, handleACpowerSupply = NULL;

	// CONSTRUCTORS FOR TestApp
	TestApp::~TestApp() {
	}

	TestApp::TestApp(CWnd* pParent) {		
	}

	
	//  RunPowerSupplyTest() - implements sequence for testing DC950 lamp voltage output
	//	while adjusting variable power supply to AC test voltages: 
	//	90, 115, 132, 254, 230, 180 volts AC.
	//
	//  The multimeter measures the lamp voltage each time 
	//  and this routine determines whether voltage is within limits.
	//  It returns NOT_DONE_YET each time through until
	//  the full sequence has completed, at which point it returns PASS.
	//	If lamp voltage is outside of limits at any point,
	//  it returns FAIL.

	int TestApp::RunPowerSupplyTest()
	{
#define NUM_TEST_VOLTAGES 6
		int supplyVoltage;
		int arrSupplyVoltage[NUM_TEST_VOLTAGES] = { 90, 115, 132, 254, 230, 180}; // Array of output AC test voltages 
		float lampVoltage, errorVoltage, expectedVoltage, controlVoltage;
		int testResult = PASS;

		if (ptrDialog->subStepNumber == 0) {
			if (!InitializeInterfaceBoard()) return FAIL;
			msDelay(100);
			if (!SetInterfaceBoardPWM(MAX_PWM)) return FAIL;
			msDelay(100);
		}

		if (ptrDialog->subStepNumber > NUM_TEST_VOLTAGES) return FAIL;
		if (ptrDialog->subStepNumber < 0) return FAIL;

		supplyVoltage = arrSupplyVoltage[ptrDialog->subStepNumber];
		if (!SetPowerSupplyVoltage(supplyVoltage)) return FAIL;

		msDelay(500);
		if (!ReadVoltage(CONTROL_VOLTAGE, &controlVoltage)) return FAIL;
		expectedVoltage = controlVoltage * (float) 4.2;
		msDelay(100);
		if (!ReadVoltage(LAMP, &lampVoltage)) return FAIL;
		errorVoltage = getAbs(lampVoltage - expectedVoltage);
		if (errorVoltage > AllowableLampVoltageError)
			testResult = FAIL;
		
		float minVoltage = expectedVoltage - AllowableLampVoltageError;
		float maxVoltage = expectedVoltage + AllowableLampVoltageError;			
		char strAddLog[MAXSTRING];
		sprintf_s(strAddLog, MAXSTRING, "AC Sweep %d volt: MIN %.2f < %.2f < MAX %.2f:", supplyVoltage, minVoltage, lampVoltage, maxVoltage);
		if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
		else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
		DisplayLog(strAddLog);			

		char strMeasuredVoltage[MAXSTRING];
		sprintf_s(strMeasuredVoltage, MAXSTRING, "%.2f", lampVoltage);
		writeTestResultsToSpreadsheet(ptrDialog->subStepNumber + SWEEP_0, testResult, strMeasuredVoltage);

		if (testResult == FAIL) return (FAIL);

		ptrDialog->subStepNumber++;
		if (ptrDialog->subStepNumber >= NUM_TEST_VOLTAGES) {
			SetPowerSupplyVoltage(120);
			InitializeInterfaceBoard();
			return (PASS);
		}
		else return NOT_DONE_YET;
	}


	// Ths routine tests three remote control features on the DC950:
	// the 5V reference voltage output, 
	// the remote pot voltage control input,
	// and the remote relay inhibit input.
	//
	// For each test, the lamp voltage is measured 
	// to ensure it is proportional to the pot input voltage. 
	// This routine runs a different
	// pot voltage from 0 to five volts each time through 
	// until the test sequence is completed.
	// 
	// The first and last voltage in the sequence are repeated:
	// 0 volts is tested twice so the 5V reference can be checked,
	// and five volts is tested twice so the relay inhibit signal
	// can be checked. 	
	//
	// The multimeter measures the lamp voltage each time 
	// and this routine determines whether voltage is within limits.
	//
	// This routine returns NOT_DONE_YET each time through until
	// the full sequence has completed, at which point it returns PASS.
	// However, if lamp voltage is outside of limits at any point,
	// the FAIL is returned.
	int TestApp::RunRemoteTests()
	{
#define NUM_TEST_VALUES 8
#define LAST_TEST (NUM_TEST_VALUES - 1)
		char strAddLog[MAXSTRING];
		int testPWM;
		// Array of PWM values to produce the desired remote pot voltage
		// PWM = 1243 yields 1 volt approximately
		// PWM = 2486 yields 2 volts
		// PWM = 3729 yields 3 volts
		// PWM = 4972 yields 4 volts
		// PWM = 6215 yields 5 volts
		int PWMvalues[NUM_TEST_VALUES] = { 0, 0, 1243, 2486, 3729, 4972, 6215, 6215 };  
		float measuredVoltage, controlVoltage, expectedVoltage, errorVoltage, maxVoltage, minVoltage;
		int testResult = PASS;

		if (ptrDialog->subStepNumber > NUM_TEST_VALUES) return FAIL;
		if (ptrDialog->subStepNumber < 0) return FAIL;
		testPWM = PWMvalues[ptrDialog->subStepNumber];

		// If this is the first Remote voltage test,
		// make sure the INHIBIT relay on the interface board is OFF,
		// and set the expected voltage to 5.0 volts
		// so that VREF can be checked
		if (ptrDialog->subStepNumber == 0) {
			SetInhibitRelay(FALSE);
		}

		msDelay(100);
		if (!SetInterfaceBoardPWM(testPWM)) return FAIL;		

		// FIRST TEST IN SEQUENCE: CHECK 5V VREF VOLTAGE OUTUPUT:
		if (ptrDialog->subStepNumber == 0) 
		{
			msDelay(500);

			expectedVoltage = (float) 5.0;
			minVoltage = expectedVoltage - AllowableVrefError;
			maxVoltage = expectedVoltage + AllowableVrefError;	

			if (!ReadVoltage(VREF, &measuredVoltage)) return FAIL;
			errorVoltage = getAbs(measuredVoltage - expectedVoltage);
			if (errorVoltage > AllowableVrefError) testResult = FAIL;
				
			sprintf_s(strAddLog, MAXSTRING, "5V Reference:      MIN %.2f < %.2f < MAX %.2f:", minVoltage, measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);			
		}
		// TEST REMOTE POT CONTROL VOLTAGES 0-5 VOLTS:
		else if (ptrDialog->subStepNumber < LAST_TEST) 
		{
			msDelay(500);
			if (!ReadVoltage(CONTROL_VOLTAGE, &controlVoltage)) return FAIL;
			expectedVoltage = controlVoltage * (float) 4.2;
			msDelay(100);
			if (!ReadVoltage(LAMP, &measuredVoltage)) return FAIL;
			errorVoltage = getAbs(measuredVoltage - expectedVoltage);
			if (errorVoltage > AllowableLampVoltageError)
				testResult = FAIL;

			minVoltage = expectedVoltage - AllowableLampVoltageError;
			maxVoltage = expectedVoltage + AllowableLampVoltageError;		
			sprintf_s(strAddLog, MAXSTRING, "Remote %d volt:      MIN %.2f < %.2f < MAX %.2f:", ptrDialog->subStepNumber-1, minVoltage, measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);			
		}
		// LAST TEST IN SEQUENCE - CHECK INHIBIT RELAY INPUT TO DC950.
		// WITH POT CONTROL INPUT AT 5 VOLTS AND INHIBIT RELAY ON,
		// LAMP SHOULD BE OFF:
		else {
			SetInhibitRelay(TRUE);
			msDelay(500);		

			expectedVoltage = (float) 0.0;
			if (!ReadVoltage(LAMP, &measuredVoltage)) return FAIL;			
			errorVoltage = getAbs(measuredVoltage - expectedVoltage);
			if (errorVoltage > AllowableLampVoltageError)
				testResult = FAIL;

			maxVoltage = expectedVoltage + AllowableLampVoltageError;
			sprintf_s(strAddLog, MAXSTRING, "Inhibit ON, lamp OFF:      %.2f volts  <   MAX %.2f:", measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);
		}


		char strMeasuredVoltage[MAXSTRING];
		sprintf_s(strMeasuredVoltage, MAXSTRING, "%.2f", measuredVoltage);
		writeTestResultsToSpreadsheet(ptrDialog->subStepNumber + VREF_TEST, testResult, strMeasuredVoltage);

		if (testResult == FAIL) return (FAIL);

		ptrDialog->subStepNumber++;
		if (ptrDialog->subStepNumber >= NUM_TEST_VALUES) {
			InitializeInterfaceBoard();
			return (PASS);
		}
		else return NOT_DONE_YET;
	}

	// SPECTROMETER TEST SEQUENCE
	int TestApp::RunSpectrometerTest(BOOL filterIsOn)
	{
		char strText[MAXSTRING];
		int testResult = NOT_DONE_YET;		

#define MAXSECONDS 60
		if (ptrDialog->subStepNumber == 0) {
			int intPWM;
			if (filterIsOn) intPWM = MAX_PWM;
			else intPWM = MAX_PWM / 10;
			if (!SetInterfaceBoardPWM(intPWM)) return FAIL;	
			msDelay(100);
			if (!SetInterfaceBoardActuatorOutput(filterIsOn)) return FAIL;
		}
		else if (ptrDialog->subStepNumber == 1)
			ActivateSpectrometer();
		else if (ptrDialog->subStepNumber == 2)
			startMeasurement();
		else if (ptrDialog->subStepNumber > MAXSECONDS) {
			stopMeasurement();
			DisplayText(3, "ERROR: Scan timeout");
			return FAIL;
		}
		else if (spectrometerDataReady()) {
			stopMeasurement();
			DisplayText(3, "SUCCESS!! SCAN COMPLETED");	
			
			getSpectrometerData (&wavelength, &minAmplitude, &peakIrradiance, &FWHM, &colorTemperature);

			testResult = PASS; // Assume success, change to FAIL if any of the following limits are exceeded:

			// SPECTROMETER TEST FOR CLOSED FILTER:
			if (filterIsOn) {
				// CENTER WAVELNGTH CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", wavelength);
				if (wavelength > MaxClosedFilterWavelength || wavelength < MinClosedFilterWavelength) {
					writeTestResultsToSpreadsheet(WAVELENGTH_CLOSED, FAIL, strText);
					testResult = FAIL;
				}
				else writeTestResultsToSpreadsheet(WAVELENGTH_CLOSED, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Wavelength filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterWavelength, wavelength, MaxClosedFilterWavelength);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// IRRADIANCE CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", peakIrradiance);
				if (peakIrradiance > MaxClosedFilterIrradiance || wavelength < MinClosedFilterIrradiance) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(IRRADIANCE_CLOSED, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(IRRADIANCE_CLOSED, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Irradiance filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterIrradiance, peakIrradiance, MaxClosedFilterIrradiance);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// MIN_AMPLITUDE_CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", minAmplitude);
				if (minAmplitude > MinClosedFilterAmplitude) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(MIN_AMPLITUDE_CLOSED, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(MIN_AMPLITUDE_CLOSED, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Min amplitude filter closed:      MIN %.2f < %.2f:", MinClosedFilterIrradiance, minAmplitude);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// FWHM_CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", FWHM);
				if (FWHM > MaxClosedFilterFWHM || FWHM < FWHM) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(FWHM_CLOSED, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(FWHM_CLOSED, PASS, strText);

				sprintf_s(strText, MAXSTRING, "FWHM filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterFWHM, peakIrradiance, MaxClosedFilterFWHM);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// COLOR TEMPERATURE CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", colorTemperature);
				if (colorTemperature > MaxClosedFilterColorTemp || colorTemperature < MinClosedFilterColorTemp) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(COLOR_TEMP_CLOSED, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(COLOR_TEMP_CLOSED, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Color Temp closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterColorTemp, colorTemperature, MaxClosedFilterColorTemp);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);
			}
			// SPECTROMETER TEST FOR FILTER OPEN:
			else {
				// IRRADIANCE OPEN:
				sprintf_s(strText, MAXSTRING, "%.2f", peakIrradiance);
				if (peakIrradiance > MaxOpenIrradiance || wavelength < MinOpenIrradiance) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(IRRADIANCE_OPEN, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(IRRADIANCE_OPEN, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Irradiance filter open:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterIrradiance, peakIrradiance, MaxClosedFilterIrradiance);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// COLOR TEMPERATURE OPEN:
				sprintf_s(strText, MAXSTRING, "%.2f", colorTemperature);
				if (colorTemperature > MaxOpenColorTemp || colorTemperature < MinOpenColorTemp) {
					testResult = FAIL;
					writeTestResultsToSpreadsheet(COLOR_TEMP_OPEN, FAIL, strText);
				}
				else writeTestResultsToSpreadsheet(COLOR_TEMP_OPEN, PASS, strText);

				sprintf_s(strText, MAXSTRING, "Color Temp open:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterColorTemp, colorTemperature, MaxClosedFilterColorTemp);
				if (testResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);
			}
		}
		else {
			sprintf_s(strText, MAXSTRING, "Scan time: %d seconds", ptrDialog->subStepNumber);
			DisplayText(3, strText);
			testResult = NOT_DONE_YET;
		}
		ptrDialog->subStepNumber++;
		return testResult;
	}


	// CHECK TTL FAULT SIGNAL - OPEN AND CLOSE ACTUATOR A FEW TIMES 
	// AND MAKE SURE TTL OUTPUT REMAINS HIGH:
	int TestApp::TTL_inputTest(){
		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED)) return (FAIL);
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_OPEN)) return (FAIL);
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED)) return (FAIL);
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_OPEN)) return (FAIL);
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED)) return (FAIL);		
		return PASS;
	}	
	  

	// EXECUTE(): THS ROUTINE IS THE MAIN FUNCTION WHICH EXECUTES ALL TESTS.
	// It is called by Testhandler() in DC950TestDialog. 
	// The global variable "stepNumber" determines which test in the sequence
	// is executed. 
	// 
	int TestApp::Execute(int stepNumber) {

		int testStatus = PASS; 			
		float fltVoltage;		
		char chrSerialNumber[MAXSTRING];
		char strTest[MAXSTRING];
		float minVoltage, maxVoltage;
		
		switch (stepNumber) {
		case 0:			// 0 Start Up
			testStatus = SystemStartUp();
			break;

		case 1:			// 1 SCAN BARCODE
			msDelay(100);			
			SetPowerSupplyVoltage(120);
			ptrDialog->m_BarcodeEditBox.GetWindowText((LPTSTR)chrSerialNumber, MAXSTRING);
			ClearLog();
			
			if (0 == strcmp(chrSerialNumber, "")) {
				testStatus = FAIL;
				break;
			}
			writeDateAndSerialNumberToSpreadsheet(chrSerialNumber);
			testStatus = PASS;
			ptrDialog->runFilterActuatorTest = (BOOL)ptrDialog->m_static_optFilter.GetCheck();
			ptrDialog->m_static_optFilter.EnableWindow(FALSE);
			ptrDialog->m_static_optStandard.EnableWindow(FALSE);
			if (ptrDialog->runFilterActuatorTest == FALSE){
				ptrDialog->m_EditBox_Test8.EnableWindow(FALSE);
				ptrDialog->m_EditBox_Test9.EnableWindow(FALSE);
				ptrDialog->m_EditBox_Test10.EnableWindow(FALSE);
			}
			else 
			{
				ptrDialog->m_EditBox_Test8.EnableWindow(TRUE);
				ptrDialog->m_EditBox_Test9.EnableWindow(TRUE);
				ptrDialog->m_EditBox_Test10.EnableWindow(TRUE);
			}
			ptrDialog->subStepNumber = 0;									
			break;

		case 2:			// 2 HI POT TEST	
			testStatus = testStep[stepNumber].Status;
			DisplayTestEditBox(HI_POT_EDIT, testStatus);
			if (testStatus == PASS) DisplayLog("Hi Pot test:    PASS\r\n");
			else DisplayLog("Hi Pot test:    FAIL\r\n");
			writeTestResultsToSpreadsheet(HI_POT, testStatus, NULL);
			break;

		case 3:			// 3 GROUND BOND TEST		
			testStatus = testStep[stepNumber].Status;
			DisplayTestEditBox(GROUND_BOND_EDIT, testStatus);
			if (testStatus == PASS) DisplayLog("Ground Bond test:    PASS\r\n");
			else DisplayLog("Ground Bond test:    FAIL\r\n");
			writeTestResultsToSpreadsheet(GROUND_BOND, testStatus, NULL);
			break;
		case 4:			// 4 POT TEST LOW = LAMP OFF	
			ReadVoltage(LAMP, &fltVoltage);
			if (fltVoltage  > AllowableLampVoltageError) testStatus = FAIL;
			else testStatus = PASS;
			DisplayTestEditBox(POT_LOW_EDIT, testStatus);

			maxVoltage = AllowableLampVoltageError;		
			sprintf_s(strTest, MAXSTRING, "Pot turned OFF: MIN 0  <  %.2f volts  <   MAX %.2f:", fltVoltage, maxVoltage);
			if (testStatus == PASS) strcat_s (strTest, MAXSTRING, "    PASS\r\n");
			else strcat_s (strTest, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strTest);			
			writeTestResultsToSpreadsheet(LAMP_OFF, testStatus, NULL);
			break;
		case 5:			// 5 POT TEST HIGH = LAMP ON
			ReadVoltage(LAMP, &fltVoltage);
			if (getAbs(fltVoltage - (float) 21.0)  > AllowableLampVoltageError) testStatus = FAIL;
			else testStatus = PASS;
			DisplayTestEditBox(POT_HIGH_EDIT, testStatus);

			minVoltage = (float) 21.0 - AllowableLampVoltageError;
			maxVoltage = (float) 21.0 + AllowableLampVoltageError;			
			sprintf_s(strTest, MAXSTRING, "Pot at FULL: MIN %.2f  <  %.2f volts  <  MAX %.2f:", minVoltage, fltVoltage, maxVoltage);
			if (testStatus == PASS) strcat_s (strTest, MAXSTRING, "    PASS\r\n");
			else strcat_s (strTest, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strTest);
			writeTestResultsToSpreadsheet(LAMP_ON, testStatus, NULL);
			break;
		case 6:			// 6 REMOTE TEST SETUP		
			break;
		case 7:			// 7 REMOTE POT TEST
			testStatus = RunRemoteTests();
			DisplayTestEditBox(REMOTE_EDIT, testStatus);								
			break;

		case 8:			// 8 AC POWER SWEEP
			testStatus = RunPowerSupplyTest();
			DisplayTestEditBox(AC_SWEEP_EDIT, testStatus);	
			DisplayTestEditBox(SIGNAL_EDIT, testStatus);
			break;

		case 9:			// 9 ACTUATOR TTL INPUT CHECK
			testStatus = TTL_inputTest();
			DisplayTestEditBox(ACTUATOR_EDIT, testStatus);
			if (testStatus == PASS) DisplayLog("TTL check:    PASS\r\n");
			else DisplayLog("TTL check:    FAIL\r\n");
			writeTestResultsToSpreadsheet(TTL_OUT, testStatus, NULL);
			break;

		case 10:			// 11 SPECTROMETER TEST - FILTER CLOSED
			testStatus = RunSpectrometerTest(FILTER_CLOSED);
			if (testStatus !=  NOT_DONE_YET) InitializeInterfaceBoard();
			DisplayTestEditBox(FILTER_ON_EDIT, testStatus);
			break;

		case 11:			// 11 SPECTROMETER TEST - FILTER OPEN
			testStatus = RunSpectrometerTest(FILTER_OPEN);
			if (testStatus !=  NOT_DONE_YET) InitializeInterfaceBoard();
			DisplayTestEditBox(FILTER_OFF_EDIT, testStatus);
			break;
			
		case 12:		// 12 TEST COMPLETE
			writeTestResultsToSpreadsheet(FINAL_RESULT, PASS, NULL);			
			SetPowerSupplyVoltage(120);
			resetTestData();
			resetDisplays();
			testStatus = PASS;
			DisplayTestEditBox(RESULT_EDIT, testStatus);
			InitializeInterfaceBoard();
			DisplayLog("UNIT PASSED\r\n");
			break;

		case 13:		// 13 UNIT FAILED
			writeTestResultsToSpreadsheet(FINAL_RESULT, FAIL, NULL);			
			SetPowerSupplyVoltage(120);
			resetTestData();
			resetDisplays();
			testStatus = FAIL;
			DisplayTestEditBox(RESULT_EDIT, testStatus);
			InitializeInterfaceBoard();
			DisplayLog("UNIT FAILED\r\n");
		default:
			break;
		}
		testStep[stepNumber].Status = testStatus;
		return testStatus;
	}
		
	
