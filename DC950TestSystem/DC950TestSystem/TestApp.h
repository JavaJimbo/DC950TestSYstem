#ifndef _TEST_ROUTINES_H
#define _TEST_ROUTINES_H
#include "stdafx.h"
#include "DC950TestDialog.h"
#include "Definitions.h"
#include <iostream>
#include <sstream>

class TestApp : public CDialog
{
public:					
	BOOL TestApp::sendReceiveSerial (int COMdevice, char *outPacket, char *inPacket, BOOL useCRC);
	void msDelay (int milliseconds);	

	BOOL openSerialPort(const char *ptrPortName, HANDLE *ptrPortHandle);
	BOOL closeSerialPort(HANDLE ptrPortHandle);
	void closeAllSerialPorts();		
	BOOL ReadSerialPort(HANDLE ptrPortHandle, char *ptrPacket);
	BOOL WriteSerialPort(HANDLE ptrPortHandle, char *ptrPacket);
	
	UINT SystemStartUp();

	BOOL InitializeTestEditBoxes();
	
	
	void resetTestData();	
	int  Execute(int stepNumber);		
	void enableBarcodeScan();
	void disableBarcodeScan();
	
	BOOL SetPowerSupplyVoltage(int voltage);	
	BOOL SetInterfaceBoardMulitplexer(int multiplexerSelect);
	BOOL ReadVoltage(int multiplexerSelect, float *ptrLampVoltage);
	float getAbs(float floatValue);
	BOOL SetInterfaceBoardPWM(int PWMvalue);
	int  RunRemoteTests();
	int  RunPowerSupplyTest();
	BOOL SetInhibitRelay(BOOL flagON);
	BOOL SetInterfaceBoardActuatorOutput(BOOL actuatorClosed);
	BOOL getSpreadSheetFileame(char *ptrExcelFilename);
	BOOL createNewSpreadsheet();	
	BOOL loadSpreadsheet();	
	BOOL writeDateAndSerialNumberToSpreadsheet(char *ptrSerialNumber); 
	BOOL writeTestResultsToSpreadsheet(int col, int testResult, char *ptrTestData);
	BOOL CloseSpreadsheet();
		
	BOOL ActivateSpectrometer();
	BOOL getSpectrometerConfiguration();
	BOOL startMeasurement();
	BOOL stopMeasurement();
	BOOL spectrometerDataReady();
	BOOL CloseSpectrometer();	
	BOOL getSpectrometerData(float *wavelength, float *minAmplitude, float *peakIrradiance, float *FWHM, float *colorTemperature);	
	BOOL getWavelengthAndIrradiance(float *peakIrradiance, float *wavelength);	
	BOOL getMinAmplitude(float *minAmplitude);
	BOOL writeSpectrometerDataToFile(const char *filename);
	int  RunSpectrometerTest(BOOL filterIsOn);

	BOOL InitializeSystem();
	BOOL InitializeHP34401();
	BOOL InitializeInterfaceBoard();
	BOOL InitializePowerSupply();
	BOOL InitializeSpectrometer();

	void ClearLog();
	int  TTL_inputTest();	
	void ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold);
	BOOL InitializeFonts();

	void write_INI_File();
	BOOL load_INI_File();

	TestStep testStep[TOTAL_STEPS];		
	CFont BigFont, SmallFont, MidFont;	
	CTestDialog *ptrDialog;

	BOOL InitializeDisplayText();
	void DisplayLog(char *newString);
	BOOL DisplayTestEditBox(int boxNumber, int passFailStatus);	
	void DisplayIntructions(int stepNumber);	
	void DisplayPassFailStatus(int stepNumber, int subStepNumber);
	void DisplayStatusBarText(int panel, LPCTSTR strText);
	void DisplaySerialComData(int comDirection, char *strData);
	void DisplayText(int lineNumber, char *strText);	
	void DisplayStepNumber(int stepNumber);			
	void resetDisplays();
	BOOL DisplayMessageBox(char *strTopLine, char *strBottomLine, int boxType);

	void writeBinaryFile(std::ostream& out, float value);
	void readBinaryFile(std::istream& in, float& value);	

	// void write(std::ostream& out, long value);	// Write binary long value
	// void read(std::istream& in, long& value);	// Read binary long value	

	// TestApp(CWnd* pParent = NULL);	// standard constructor
	TestApp::TestApp(CWnd* pParent = NULL);
	~TestApp();
};

#endif

