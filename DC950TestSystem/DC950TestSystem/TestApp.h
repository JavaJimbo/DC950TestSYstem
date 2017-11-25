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
	BOOL TestApp::sendReceiveSerial (int COMdevice, CTestDialog *ptrDialog, char *outPacket, char *inPacket, BOOL useCRC);
	void msDelay (int milliseconds);	

	BOOL openSerialPort(const char *ptrPortName, HANDLE *ptrPortHandle);
	BOOL closeSerialPort(HANDLE ptrPortHandle);
	void closeAllSerialPorts();		
	BOOL ReadSerialPort(HANDLE ptrPortHandle, char *ptrPacket);
	BOOL WriteSerialPort(HANDLE ptrPortHandle, char *ptrPacket);
	
	UINT SystemStartUp(CTestDialog *ptrDialog);

	BOOL InitializeTestEditBoxes(CTestDialog *ptrDialog);
	
	
	void resetTestData();	
	int  Execute(int stepNumber, CTestDialog *ptrDialog);		
	void enableBarcodeScan(CTestDialog *ptrDialog);
	void disableBarcodeScan(CTestDialog *ptrDialog);
	
	BOOL SetPowerSupplyVoltage(CTestDialog *ptrDialog, int voltage);	
	BOOL SetInterfaceBoardMulitplexer(CTestDialog *ptrDialog, int multiplexerSelect);
	BOOL ReadVoltage(CTestDialog *ptrDialog, int multiplexerSelect, float *ptrLampVoltage);
	float getAbs(float floatValue);
	BOOL SetInterfaceBoardPWM(CTestDialog *ptrDialog, int PWMvalue);
	int  RunRemoteTests(CTestDialog *ptrDialog);
	int  RunPowerSupplyTest(CTestDialog *ptrDialog);
	BOOL SetInhibitRelay(CTestDialog *ptrDialog, BOOL flagON);
	BOOL SetInterfaceBoardActuatorOutput(CTestDialog *ptrDialog, BOOL actuatorClosed);
	BOOL getSpreadSheetFileame(char *ptrExcelFilename);
	BOOL createNewSpreadsheet();	
	BOOL loadSpreadsheet();	
	BOOL writeDateAndSerialNumberToSpreadsheet(char *ptrSerialNumber); 
	BOOL writeTestResultsToSpreadsheet(int col, int testResult, char *ptrTestData);
	BOOL CloseSpreadsheet();
		
	BOOL ActivateSpectrometer(CTestDialog *ptrDialog);
	BOOL getSpectrometerConfiguration(CTestDialog *ptrDialog);
	BOOL startMeasurement();
	BOOL stopMeasurement();
	BOOL spectrometerDataReady();
	BOOL CloseSpectrometer();	
	BOOL getSpectrometerData(CTestDialog *ptrDialog, float *wavelength, float *minAmplitude, float *peakIrradiance, float *FWHM, float *colorTemperature);	
	BOOL getWavelengthAndIrradiance(float *peakIrradiance, float *wavelength);	
	BOOL getMinAmplitude(float *minAmplitude);
	BOOL writeSpectrometerDataToFile(const char *filename);
	int  RunSpectrometerTest(BOOL filterIsOn, CTestDialog *ptrDialog);

	BOOL InitializeSystem(CTestDialog *ptrDialog);
	BOOL InitializeHP34401(CTestDialog *ptrDialog);
	BOOL InitializeInterfaceBoard(CTestDialog *ptrDialog);
	BOOL InitializePowerSupply(CTestDialog *ptrDialog);
	BOOL InitializeSpectrometer(CTestDialog *ptrDialog);

	void ClearLog(CTestDialog *ptrDialog);
	int  TTL_inputTest(CTestDialog *ptrDialog);	
	void ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold);
	BOOL InitializeFonts();

	void write_INI_File();
	BOOL load_INI_File();

	TestStep testStep[TOTAL_STEPS];		
	CFont BigFont, SmallFont, MidFont;	
	
	//const char *portNameInterfaceBoard = "COM8";
	//const char *portNameMultiMeter = "\\\\.\\COM3";
	//const char *portNameACpowerSupply = "\\\\.\\COM19";
	// CTestDialog *ptrDialog;
			
	//const char *portNameInterfaceBoard = "COM3";
	//const char *portNameACpowerSupply = "COM4";
	//const char *portNameMultiMeter = "COM6";

	BOOL InitializeDisplayText();
	void DisplayLog(char *newString, CTestDialog *ptrDialog);
	BOOL DisplayTestEditBox(CTestDialog *ptrDialog, int boxNumber, int passFailStatus);	
	void DisplayIntructions(int stepNumber, CTestDialog *ptrDialog);	
	void DisplayPassFailStatus(int stepNumber, int subStepNumber, CTestDialog *ptrDialog);
	void DisplayStatusBarText(CTestDialog *ptrDialog, int panel, LPCTSTR strText);
	void DisplaySerialComData(CTestDialog *ptrDialog, int comDirection, char *strData);
	void DisplayText(CTestDialog *ptrDialog, int lineNumber, char *strText);	
	void DisplayStepNumber(CTestDialog *ptrDialog, int stepNumber);			
	void resetDisplays(CTestDialog *ptrDialog);
	BOOL DisplayMessageBox(char *strTopLine, char *strBottomLine, int boxType);

	void writeBinaryFile(std::ostream& out, float value);
	void readBinaryFile(std::istream& in, float& value);	

	// void write(std::ostream& out, long value);	// Write binary long value
	// void read(std::istream& in, long& value);	// Read binary long value	

	TestApp(CWnd* pParent = NULL);	// standard constructor
	~TestApp();
};

#endif

