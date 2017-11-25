/*  Project: DC950TestSystem
*   DC950TestDialog.cpp - implementation file for main dialog box 
*
 */
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include "TestApp.h"
#include "Definitions.h"
#include <windows.h>
#include <iostream>
#include <sstream>
#include <string.h>
#include "avaspec.h"
#include "BasicExcel.hpp"
#include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	

using namespace WinCompFiles;
using namespace YExcel;
using namespace ExcelFormat;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

TestApp MyTestApp;	
CString m_ToAdd;
long testCount = 16;


// Dialog deconstructor: shut down system 
CTestDialog::~CTestDialog() {	
	MyTestApp.closeAllSerialPorts();
	MyTestApp.CloseSpreadsheet();	
}

// Dialog constructor
CTestDialog::CTestDialog(CWnd* pParent /*=NULL*/)	: CDialog(CTestDialog::IDD, pParent)	// , bPortOpened(FALSE)
{
	stepNumber = 0;
	subStepNumber = 0;
	runFilterActuatorTest = FALSE;	
	m_timerHandle = NULL;	
	m_ToAdd = "This is a test"; // _T("");
}



void CTestDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	
	DDX_Control(pDX, IDC_STATIC_PASSFAIL_STATUS, m_static_TestResult);
	
	DDX_Control(pDX, IDC_STATIC_LINE1, m_static_Line1);
	DDX_Control(pDX, IDC_STATIC_LINE2, m_static_Line2);
	DDX_Control(pDX, IDC_STATIC_LINE3, m_static_Line3);
	DDX_Control(pDX, IDC_STATIC_LINE4, m_static_Line4);
	DDX_Control(pDX, IDC_STATIC_LINE5, m_static_Line5);	
	DDX_Control(pDX, IDC_STATIC_LINE6, m_static_Line6);
	DDX_Control(pDX, IDC_STATIC_LINE7, m_static_DataOut);
	DDX_Control(pDX, IDC_STATIC_LINE8, m_static_DataIn);
	DDX_Control(pDX, IDC_STATIC_RESULT, m_static_TestName);
	DDX_Control(pDX, IDC_STATIC_COM_OUT, m_static_ComOut);
	DDX_Control(pDX, IDC_STATIC_COM_IN, m_static_ComIn);
		
	DDX_Control(pDX, IDC_BUTTON_ENTER, m_static_ButtonEnter); 
	DDX_Control(pDX, IDC_BUTTON_PASS,  m_static_ButtonPass);
	DDX_Control(pDX, IDC_BUTTON_ADMIN,  m_static_ButtonAdmin);
	
	DDX_Control(pDX, IDC_BUTTON_FAIL, m_static_ButtonFail);
	DDX_Control(pDX, IDC_BUTTON_TEST, m_static_ButtonTest);	
	DDX_Control(pDX, IDC_BUTTON_PREVIOUS, m_static_ButtonPrevious);
	DDX_Control(pDX, IDC_BUTTON_HALT, m_static_ButtonHalt);
	DDX_Control(pDX, IDC_EDIT1, m_BarcodeEditBox);

	DDX_Control(pDX, IDC_EDIT_TEST1, m_EditBox_Test1);
	DDX_Control(pDX, IDC_EDIT_TEST2, m_EditBox_Test2);
	DDX_Control(pDX, IDC_EDIT_TEST3, m_EditBox_Test3);
	DDX_Control(pDX, IDC_EDIT_TEST4, m_EditBox_Test4);
	DDX_Control(pDX, IDC_EDIT_TEST5, m_EditBox_Test5);
	DDX_Control(pDX, IDC_EDIT_TEST6, m_EditBox_Test6);
	DDX_Control(pDX, IDC_EDIT_TEST7, m_EditBox_Test7);
	DDX_Control(pDX, IDC_EDIT_TEST8, m_EditBox_Test8);
	DDX_Control(pDX, IDC_EDIT_TEST10, m_EditBox_Test9);
	DDX_Control(pDX, IDC_EDIT_TEST11, m_EditBox_Test10);

	DDX_Control(pDX, IDC_EDIT_LOG, m_EditBox_Log);

	DDX_Control(pDX, IDC_STATIC_SN, m_static_BarcodeLabel);
	
	DDX_Control(pDX, IDC_RADIO_STD, m_static_optStandard);
	DDX_Control(pDX, IDC_RADIO_FILTER, m_static_optFilter);
	DDX_Control(pDX, IDC_STATIC_MODELGROUP, m_static_modelGroup);
}


BEGIN_MESSAGE_MAP(CTestDialog, CDialog)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(IDC_BUTTON_TEST, &CTestDialog::OnBtnClickedTest)
	ON_BN_CLICKED(IDC_BUTTON_ENTER, &CTestDialog::OnBtnClickedEnter)
	ON_BN_CLICKED(IDC_BUTTON_HALT, &CTestDialog::OnBtnClickedHalt)
	ON_BN_CLICKED(IDC_BUTTON_PREVIOUS, &CTestDialog::OnBtnClickedPrevious)
	ON_BN_CLICKED(IDC_BUTTON_PASS, &CTestDialog::OnBtnClickedPass)
	ON_BN_CLICKED(IDC_BUTTON_FAIL, &CTestDialog::OnBtnClickedFail)
	ON_BN_CLICKED(IDC_RADIO_STD, &CTestDialog::OnBtnClickedRadioSTD)
	ON_BN_CLICKED(IDC_RADIO_FILTER, &CTestDialog::OnBtnClickedRadioFilter)	
	ON_STN_CLICKED(IDC_STATIC_LINE4, &CTestDialog::OnStnClickedStaticLine4)
	ON_BN_CLICKED(IDC_BUTTON_ADMIN, &CTestDialog::OnBnClickedButtonAdmin)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

void CTestDialog::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialog::OnSysCommand(nID, lParam);
}


// PREVIOUS BUTTON
void CTestDialog::OnBtnClickedPrevious()
{
	int  stepType;
	do {
		if (stepNumber <= BARCODE_SCAN) break;
		stepNumber--;
		if (stepNumber == REMOTE_TEST) break;
		
		stepType = MyTestApp.testStep[stepNumber].stepType;
	} while (stepType == AUTO);

	if (stepNumber == 1) MyTestApp.resetDisplays(this); 
	
	MyTestApp.DisplayStepNumber(this, stepNumber);
	MyTestApp.DisplayIntructions(stepNumber, this);
	MyTestApp.DisplayPassFailStatus(stepNumber, 0, this);
}







BOOL CTestDialog::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_RETURN || pMsg->wParam == VK_ESCAPE)
		{
			return TRUE;                // Do not process further
		}
	}

	return CWnd::PreTranslateMessage(pMsg);
}


// CTestDialog message handlers

BOOL CTestDialog::OnInitDialog()
{	
	CDialog::OnInitDialog();	
	
	enableAdmin = FALSE;
	pEdit = (CEdit*) GetDlgItem(IDC_EDIT_LOG);
	MyTestApp.InitializeFonts();
	
	CWnd* pCtrlWnd = GetDlgItem(IDC_STATIC_LINE1);
	pCtrlWnd->SetFont(&MyTestApp.BigFont, TRUE);

	pCtrlWnd = GetDlgItem(IDC_STATIC_PASSFAIL_STATUS);
	pCtrlWnd->SetFont(&MyTestApp.MidFont, TRUE);

	pCtrlWnd = GetDlgItem(IDC_STATIC_LINE6);
	pCtrlWnd->SetFont(&MyTestApp.SmallFont, TRUE);
	
	stepNumber = 0;
	subStepNumber = 0;	
	MyTestApp.InitializeDisplayText();
	MyTestApp.DisplayIntructions(stepNumber, this);
	
	// MyTestApp.DisplayStepNumber(this, stepNumber);
	m_static_optStandard.SetCheck(TRUE);
	m_static_optFilter.SetCheck(FALSE);

	// MyTestApp.ptrDialog = this;
	
	m_EditBox_Log.SetLimitText(MAXLOG); 	
	MyTestApp.InitializeTestEditBoxes(this);
	MyTestApp.load_INI_File();
	// m_DialogBackground.SetBackColor(RGB(0,255,255));
	
	MyTestApp.loadSpreadsheet(); 

	return TRUE;  // return TRUE  unless you set the focus to a control
}


// HALT TIMER
void CTestDialog::OnBtnClickedHalt()
{
	StopTimer();	
}


void CTestDialog::OnBtnClickedRadioSTD()
{
	MyTestApp.enableBarcodeScan(this);
}

void CTestDialog::OnBtnClickedRadioFilter()
{
	MyTestApp.enableBarcodeScan(this);
}

void CTestDialog::OnBtnClickedEnter() {
	Testhandler();
}



void CTestDialog::OnBtnClickedPass()
{
	MyTestApp.testStep[stepNumber].Status = PASS;
	Testhandler();
}

void CTestDialog::OnBtnClickedFail()
{
	MyTestApp.testStep[stepNumber].Status = FAIL;
	Testhandler();
}


void CTestDialog::Testhandler()
{
		// Execute test step and display result:
		int testResult = MyTestApp.Execute(stepNumber, this);
		MyTestApp.DisplayPassFailStatus(stepNumber, subStepNumber, this);

		if (testResult == PASS || testResult == FAIL) {
			subStepNumber = 0;

			int stepType = MyTestApp.testStep[stepNumber].stepType;

			// Step #0 checks RS232 communication.
			// If system fails this step,
			// the interface board power probably isn't on.
			// For this case, quit routine here:
			if (stepNumber == 0 && testResult != PASS) return;

			// Otherwise if a unit has completed full test sequence,
			// set step back to 1 for testing next unit:
			else if (stepNumber == FINAL_PASS || stepNumber == FINAL_FAIL)
				stepNumber = 1;

			// If unit passedd, then advance to next step.
			// If unit isn't the filter actuator model, 
			// then skip the ACTUATOR and SPECTROMETER tests and go to FINAL PASS:
			else if (testResult == PASS) {
				stepNumber++; 
				if (runFilterActuatorTest == FALSE){
					if (stepNumber == END_STANDARD_UNIT_TESTS) 
						stepNumber = FINAL_PASS;
				}
			}
			// Otherwise if unit failed, stop timer (it it's running)
			// then query user to retry or quit testing unit:
			else {
				StopTimer();
				if (stepNumber == 1) MyTestApp.DisplayMessageBox("BARCODE ERROR", "Please scan label again", 1);
				else {
					BOOL retry = MyTestApp.DisplayMessageBox("UNIT FAILED TEST", "Retry test or Cancel?", 2);
					// If user doesn't wish to retry any failed test, jump to FAIL TEST:
					if (!retry) stepNumber = FINAL_FAIL;
				}
			}
			// If this is the start of an automated sequence
			// or if sequence is already in progress, start timer.
			// If next step is manual, make sure timer is off:
			if (MyTestApp.testStep[stepNumber].stepType == AUTO) StartTimer();
			else StopTimer();
		}

		// Now display next step:
		MyTestApp.DisplayIntructions(stepNumber, this);
		MyTestApp.DisplayStepNumber(this, stepNumber);		

		if (stepNumber == 1) MyTestApp.resetDisplays(this);// MyTestApp.enableBarcodeScan(this);
		else MyTestApp.disableBarcodeScan(this);

		// Now wait for user to hit ENTER
		// or timer to call this routine and execute new step
		return;   
}





void CTestDialog::OnStnClickedStaticLine4()
{
	// TODO: Add your control notification handler code here

}


void CTestDialog::OnBtnClickedTest() {
	
}


void CTestDialog::OnBack() 
{
	// call color dialog and change background color
	CColorDialog dlg;
	if (dlg.DoModal() == IDOK)
		m_EditBox_Test1.SetBackColor(dlg.GetColor());
	
}

void CTestDialog::OnText() 
{
	// call color dialog and change text color
	CColorDialog dlg;
	if (dlg.DoModal() == IDOK)
		m_EditBox_Test1.SetTextColor(dlg.GetColor());
}


void CTestDialog::OnBnClickedButtonAdmin()
{
	if (enableAdmin) m_Admin.DoModal();
	else {
		CString s = "";
		m_PED.DoModal(&s);	
		if(s != "Setra")	
			AfxMessageBox("Incorrect password entered!\r\nPlease see system adminstrator for access to administrator utilities.");	
		else
		{
			enableAdmin = TRUE;
			m_Admin.DoModal();
		}
	}
}

void CTestDialog::StopTimer() {
	if (m_timerHandle != NULL) {
		DeleteTimerQueueTimer(NULL, m_timerHandle, NULL);
		m_timerHandle = NULL;
	}
}

void CALLBACK TimerProc(CTestDialog *lpParametar, BOOLEAN TimerOrWaitFired)
{
	// This is used only to call QueueTimerHandler
	// Typically, this function is static member of CTimersDlg	
	CTestDialog *ptrTimer;
	ptrTimer = lpParametar;
	ptrTimer->timerHandler();			
}

void CTestDialog::StartTimer() {
	if (m_timerHandle == NULL) {
		DWORD elapsedTime = 3000;
		BOOL success = ::CreateTimerQueueTimer(  // Create new timer
			&m_timerHandle,
			NULL,
			(WAITORTIMERCALLBACK)TimerProc,
			this,
			0,
			elapsedTime,
			WT_EXECUTEINTIMERTHREAD);
	}
}

void CTestDialog::timerHandler() {
	Testhandler();
}

