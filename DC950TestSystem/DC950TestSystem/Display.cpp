/* Display.cpp - Routines for formatting and displaying text on main dialog box.
 * Written in Visual C++ for DC950 test system
 *
 * 10-28-17: Testhandler() fully implemented, test sequence flow works well.
 * 11-1-17: All tests up and running except spectrometer. Spreadsheet not done yet.
 * Comment out power supply commands
 * 
 */

// NOTE: INCUDES MUST BE IN THIS ORDER!!!
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

CReadOnlyEdit	*arrayEditBoxTest[NUM_TEST_EDIT_BOXES];

char *arrTestTitles[] = {"Hi-Pot",         "Ground Bond",     "Pot at zero",  "Pot at Full",	"Remote Lamp",  "AC Sweep",     "Signal Check",   "Actuator",     "Filter ON",     "Filter OFF",      "         "};
char arrTestLog[MAXLOG] = {'\0'};

void TestApp::ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold) {
	if (flgBold) {
		ptrFont.CreateFont(
			fontHeight,
			fontWidth,
			0,
			FW_BOLD,
			FW_DONTCARE,
			FALSE,
			FALSE,
			FALSE,
			DEFAULT_CHARSET,
			OUT_DEFAULT_PRECIS,
			CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY,
			DEFAULT_PITCH,
			NULL
		);
	}
	else
	{
		ptrFont.CreateFont(
			fontHeight,
			fontWidth,
			0,
			FW_NORMAL,
			FW_DONTCARE,
			FALSE,
			FALSE,
			FALSE,
			DEFAULT_CHARSET,
			OUT_DEFAULT_PRECIS,
			CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY,
			DEFAULT_PITCH,
			NULL
		);
	}
}

	BOOL TestApp::InitializeFonts(){
		ConfigureFont(BigFont, 24, 12, TRUE);
		ConfigureFont(MidFont, 20, 8, TRUE);
		ConfigureFont(SmallFont, 15, 5, TRUE);
		return TRUE;
	}

	BOOL TestApp::DisplayTestEditBox(CTestDialog *ptrDialog, int boxNumber, int passFailStatus){
		char TestEditBoxText[MAXSTRING];

		if (boxNumber >= NUM_TEST_EDIT_BOXES) return FALSE;

		strcpy_s (TestEditBoxText, MAXSTRING, arrTestTitles[boxNumber]);

		if (passFailStatus == NOT_DONE_YET){
			arrayEditBoxTest[boxNumber]->SetBackColor(BLACK);
			arrayEditBoxTest[boxNumber]->SetTextColor(YELLOW);		

		}

		else if (passFailStatus == PASS){
			arrayEditBoxTest[boxNumber]->SetBackColor(GREEN);
			arrayEditBoxTest[boxNumber]->SetTextColor(BLACK);		
			strcat_s (TestEditBoxText, MAXSTRING, ":  PASS");
		}

		else if (passFailStatus == FAIL){
			arrayEditBoxTest[boxNumber]->SetBackColor(RED);
			arrayEditBoxTest[boxNumber]->SetTextColor(BLACK);	
			strcat_s (TestEditBoxText, MAXSTRING, ":  FAIL");
		}

		arrayEditBoxTest[boxNumber]->SetFont(&BigFont, TRUE);
		arrayEditBoxTest[boxNumber]->SetWindowText(TestEditBoxText);
		return TRUE;
	}

	BOOL TestApp::InitializeTestEditBoxes(CTestDialog *ptrDialog)
	{
		arrayEditBoxTest[0] = &ptrDialog->m_EditBox_Test1;
		arrayEditBoxTest[1] = &ptrDialog->m_EditBox_Test2;
		arrayEditBoxTest[2] = &ptrDialog->m_EditBox_Test3;
		arrayEditBoxTest[3] = &ptrDialog->m_EditBox_Test4;
		arrayEditBoxTest[4] = &ptrDialog->m_EditBox_Test5;
		arrayEditBoxTest[5] = &ptrDialog->m_EditBox_Test6;
		arrayEditBoxTest[6] = &ptrDialog->m_EditBox_Test7;
		arrayEditBoxTest[7] = &ptrDialog->m_EditBox_Test8;	
		arrayEditBoxTest[8] = &ptrDialog->m_EditBox_Test9;		
		arrayEditBoxTest[9] = &ptrDialog->m_EditBox_Test10;		
		return TRUE;
	}

	void TestApp::resetDisplays(CTestDialog *ptrDialog) {
		ptrDialog->m_BarcodeEditBox.SetWindowText((LPCTSTR)"");
		ptrDialog->m_static_TestResult.SetWindowText((LPCTSTR)"");
		ptrDialog->m_static_TestName.SetWindowText((LPCTSTR)"New Unit");
		ptrDialog->m_static_ComOut.SetWindowText((LPCTSTR)"Serial Com Out");
		ptrDialog->m_static_ComIn.SetWindowText((LPCTSTR)"Serial Com In");
		ptrDialog->m_static_Line6.SetWindowText((LPCTSTR)"");
		ptrDialog->m_BarcodeEditBox.ShowWindow(TRUE);
		ptrDialog->m_static_BarcodeLabel.ShowWindow(TRUE);
		ptrDialog->m_static_DataIn.SetWindowText((LPCTSTR)"");
		ptrDialog->m_static_DataOut.SetWindowText((LPCTSTR)"");	
		ptrDialog->m_static_optFilter.EnableWindow(TRUE);
		ptrDialog->m_static_optStandard.EnableWindow(TRUE);
		ptrDialog->m_static_modelGroup.EnableWindow(TRUE);
		ptrDialog->m_static_BarcodeLabel.ShowWindow(TRUE);
		if (ptrDialog->enableAdmin == FALSE) ptrDialog->m_static_ButtonAdmin.ShowWindow(FALSE);
		enableBarcodeScan(ptrDialog);
		InitializeDisplayText();

		ptrDialog->m_EditBox_Test1.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test2.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test3.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test4.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test5.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test6.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test7.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test8.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test9.EnableWindow(TRUE);
		ptrDialog->m_EditBox_Test10.EnableWindow(TRUE);

		for (int i = 0; i < NUM_TEST_EDIT_BOXES; i++)  DisplayTestEditBox(ptrDialog, i, NOT_DONE_YET);
	}

	BOOL TestApp::InitializeDisplayText() {	
		int i = 0;

		// 0 PRESET		
		testStep[i].testName = NULL;
		testStep[i].lineOne = "Test system preparation:";
		testStep[i].lineTwo = "Before beginning, make sure all equipment";
		testStep[i].lineThree = "is plugged in and turned ON,";
		testStep[i].lineFour = "and INTERFACE BOX is ON.";
		testStep[i].lineFive = "Then click READY to begin.";
		testStep[i].lineSix = NULL;
		testStep[i].stepType = 0;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;		
		i++;

		// 1 SCAN BARCODE
		testStep[i].testName = "Barcode";
		testStep[i].lineOne = "Scan barcode and select model";
		testStep[i].lineTwo = "Set switch on DC950 to LOCAL";
		testStep[i].lineThree = "Plug it into power supply";
		testStep[i].lineFour = "Flip DC950 power switch to ON";
		testStep[i].lineFive = "Click ENTER to begin tests.";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;	
		testStep[i].dataColNum = 0;
		i++;

		// 2 HI POT TEST
		testStep[i].testName = testStep[i].lineOne = "Hi-Pot";
		testStep[i].lineTwo = "Connect tester, then run test.";
		testStep[i].lineThree = "Click PASS or FAIL";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = TRUE;
		testStep[i].enableFAIL = TRUE;		
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 3 GROUND BOND TEST
		testStep[i].testName = testStep[i].lineOne = "Ground Bond";
		testStep[i].lineTwo = "Connect tester, then run test.";
		testStep[i].lineThree = "Click PASS or FAIL";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = TRUE;
		testStep[i].enableFAIL = TRUE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 4 LAMP OFF
		testStep[i].testName = testStep[i].lineOne = "Lamp OFF";
		testStep[i].lineTwo = "Connect leads to lamp.";
		testStep[i].lineThree = "Set switch to LOCAL";
		testStep[i].lineFour = "Turn POT to zero, turn ON the DC950";
		testStep[i].lineFive = "Press ENTER to run test";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 5 LAMP ON
		testStep[i].testName = testStep[i].lineOne = "Lamp full ON";
		testStep[i].lineTwo = "Turn POT to MAX power";
		testStep[i].lineThree = "Press ENTER to run test";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 6 REMOTE SETUP
		testStep[i].testName = "Remote Control";
		testStep[i].lineOne = "Remote test setup:";
		testStep[i].lineTwo = "Turn pot back to zero";
		testStep[i].lineThree = "Switch DC950 to REMOTE";
		testStep[i].lineFour = "Plug in DB9 connectors";
		testStep[i].lineFive = "Press ENTER when ready";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL; //  SUBSTEP;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;
		
		// 7 REMOTE POT
		testStep[i].testName = "Remote Pot";
		testStep[i].lineOne = "Running Remote test";
		testStep[i].lineTwo = "Please wait...";
		testStep[i].lineThree = "";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;
		i++;

		// 8 AC SWEEP
		testStep[i].testName = "AC Sweep";
		testStep[i].lineOne = "Running AC Power Sweep";
		testStep[i].lineTwo = "Please wait...";
		testStep[i].lineThree = "";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 9 TTL FEEDBACK CHECK
		testStep[i].testName = "Actuator Test";
		testStep[i].lineOne = "Checking TTL input";
		testStep[i].lineTwo = "Please wait...";
		testStep[i].lineThree = "";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;
		i++;

		// 9 SPECTROMETER CLOSED FILTER
		testStep[i].testName = "Spectrometer Filter Closed";
		testStep[i].lineOne = "Running Spectrometer test";
		testStep[i].lineTwo = "Please wait...";
		testStep[i].lineThree = "";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;
		i++;

		// 10 SPECTROMETER OPEN FILTER
		testStep[i].testName = "Spectrometer Filter Open";
		testStep[i].lineOne = "Running Spectrometer test";
		testStep[i].lineTwo = "Please wait...";
		testStep[i].lineThree = "";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;
		i++;

		// 11 TEST COMPLETE PASSED
		testStep[i].testName = "Spectrometer Filter Open";
		testStep[i].lineOne = "All tests PASSED";
		testStep[i].lineTwo = "Turn off DC950";
		testStep[i].lineThree = "Disconnect test leads";
		testStep[i].lineFour = "Click ENTER to test next unit";
		testStep[i].lineFive = "";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].dataColNum = 0;
		i++;

		// 12 TEST COMPLETE FAILED	
		testStep[i].testName = NULL;
		testStep[i].lineOne = "UNIT FAILED";
		testStep[i].lineTwo = "Turn off DC950";
		testStep[i].lineThree = "Disconnect test leads";
		testStep[i].lineFour = "Attach FAIL tag";
		testStep[i].lineFive = "Press ENTER to test new unit.";
		testStep[i].lineSix = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableHALT = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].dataColNum = 0;
		return TRUE;
	}
	
	void TestApp::DisplaySerialComData(CTestDialog *ptrDialog, int comDirection, char *strData) {
		if (comDirection == DATAIN) ptrDialog->m_static_DataIn.SetWindowText((LPCTSTR)(LPCTSTR) strData);
		else ptrDialog->m_static_DataOut.SetWindowText((LPCTSTR)(LPCTSTR)strData);
	}

	// Displays user instructions for each test, and enables/disables buttons as needed		
	void TestApp::DisplayIntructions(int stepNumber, CTestDialog *ptrDialog){	
		char *ptrText = NULL;

		// Once STARTUP is done, change "READY" to "ENTER" text on button:
		if (stepNumber == 1) ptrDialog->m_static_ButtonEnter.SetWindowText((LPCTSTR)"ENTER"); 

		// Display user instructions on five text lines:
		ptrText = testStep[stepNumber].lineOne;
		if (ptrText!= NULL) ptrDialog->m_static_Line1.SetWindowText((LPCTSTR)ptrText);

		ptrText = testStep[stepNumber].lineTwo;
		if (ptrText != NULL) ptrDialog->m_static_Line2.SetWindowText((LPCTSTR)ptrText);

		ptrText = testStep[stepNumber].lineThree;
		if (ptrText != NULL) ptrDialog->m_static_Line3.SetWindowText((LPCTSTR)ptrText);

		ptrText = testStep[stepNumber].lineFour;
		if (ptrText != NULL) ptrDialog->m_static_Line4.SetWindowText((LPCTSTR)ptrText);

		ptrText = testStep[stepNumber].lineFive;
		if (ptrText != NULL) ptrDialog->m_static_Line5.SetWindowText((LPCTSTR)ptrText);

		// Enable or disable buttons depending on which step will execute next:
		ptrDialog->m_static_ButtonPass.EnableWindow(testStep[stepNumber].enablePASS);
		ptrDialog->m_static_ButtonFail.EnableWindow(testStep[stepNumber].enableFAIL);
		ptrDialog->m_static_ButtonEnter.EnableWindow(testStep[stepNumber].enableENTER);
		ptrDialog->m_static_ButtonPrevious.EnableWindow(testStep[stepNumber].enablePREVIOUS);
		ptrDialog->m_static_ButtonHalt.EnableWindow(testStep[stepNumber].enableHALT);
	}
		
	void TestApp::DisplayText(CTestDialog *ptrDialog, int lineNumber, char *strText) {
		switch (lineNumber) {
		case 1:
			ptrDialog->m_static_Line1.SetWindowText((LPCTSTR)strText);
			break;
		case 2:
			ptrDialog->m_static_Line2.SetWindowText((LPCTSTR)strText);
			break;
		case 3:
			ptrDialog->m_static_Line3.SetWindowText((LPCTSTR)strText);
			break;
		case 4:
			ptrDialog->m_static_Line4.SetWindowText((LPCTSTR)strText);
			break;
		case 5:
			ptrDialog->m_static_Line5.SetWindowText((LPCTSTR)strText);
			break;
		case 6:
			ptrDialog->m_static_Line6.SetWindowText((LPCTSTR)strText);
			break;
		default:
			break;
		}		
	}

		
	BOOL TestApp::DisplayMessageBox(char *strTopLine, char *strBottomLine, int boxType)
	{
		BOOL tryAgain = false;
		int msgBoxID;

		if (boxType == 3) msgBoxID = MessageBox((LPCTSTR)strBottomLine, (LPCTSTR)strTopLine, MB_ICONWARNING | MB_CANCELTRYCONTINUE | MB_DEFBUTTON2);
		else if (boxType == 2) msgBoxID = MessageBox((LPCTSTR)strBottomLine, (LPCTSTR)strTopLine, MB_ICONWARNING | MB_RETRYCANCEL | MB_DEFBUTTON2);
		else msgBoxID = MessageBox((LPCTSTR)strBottomLine, (LPCTSTR)strTopLine, MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);

		switch (msgBoxID)
		{
		case IDTRYAGAIN:
		case IDRETRY:
			tryAgain = true;
			break;
		case IDCANCEL:
		case IDOK:
		case IDCONTINUE:
		default:
			tryAgain = false;
			break;
		}
		return tryAgain;
	}
		
	void TestApp::enableBarcodeScan(CTestDialog *ptrDialog) {		
		ptrDialog->m_BarcodeEditBox.SetFocus();
		ptrDialog->m_BarcodeEditBox.ShowCaret();
		ptrDialog->m_BarcodeEditBox.SendMessage(EM_SETREADONLY, 0, 0);
	}

	void TestApp::disableBarcodeScan(CTestDialog *ptrDialog) {
		ptrDialog->m_BarcodeEditBox.SendMessage(EM_SETREADONLY, 1, 0);
	}

	void TestApp::DisplayStepNumber(CTestDialog *ptrDialog, int stepNumber) {
		char strStepNumber[16];
		sprintf_s(strStepNumber, "Step #%d", stepNumber);
		ptrDialog->m_static_Line6.SetWindowText((LPCTSTR)strStepNumber);
	}

void TestApp::DisplayPassFailStatus(int stepNumber, int subStepNumber, CTestDialog *ptrDialog) {
		if (stepNumber > STARTUP && stepNumber < FINAL_PASS) {
			if (testStep[stepNumber].Status == PASS) {
				ptrDialog->m_static_TestName.SetWindowText((LPCTSTR)testStep[stepNumber].testName);
				ptrDialog->m_static_TestResult.SetWindowText((LPCTSTR)"PASS");
			}
			else if (testStep[stepNumber].Status == FAIL) {
				ptrDialog->m_static_TestName.SetWindowText((LPCTSTR)testStep[stepNumber].testName);
				ptrDialog->m_static_TestResult.SetWindowText((LPCTSTR)"FAIL");
			}
			else if (testStep[stepNumber].Status == NOT_DONE_YET) {
				if (!subStepNumber) {
					ptrDialog->m_static_TestName.SetWindowText((LPCTSTR)testStep[stepNumber].testName);
					ptrDialog->m_static_TestResult.SetWindowText((LPCTSTR)"PLEASE WAIT");
				}
				else {
					ptrDialog->m_static_TestName.SetWindowText((LPCTSTR)testStep[stepNumber].testName);
					char strSubsStep[32];
					sprintf_s(strSubsStep, "Substep #%d", subStepNumber);
					ptrDialog->m_static_TestResult.SetWindowText((LPCTSTR)strSubsStep);
				}
			}
		}
	}

	void TestApp::ClearLog(CTestDialog *ptrDialog){
		arrTestLog[0] = '\0';
		ptrDialog->m_EditBox_Log.SetWindowText(arrTestLog);
		ptrDialog->pEdit->LineScroll (ptrDialog->pEdit->GetLineCount());
	}
	

	void TestApp::DisplayLog(char *newString, CTestDialog *ptrDialog){		
		strcat_s(arrTestLog, MAXLOG, newString);
		ptrDialog->m_EditBox_Log.SetWindowText(arrTestLog);
		ptrDialog->pEdit->LineScroll (ptrDialog->pEdit->GetLineCount());			
	}

