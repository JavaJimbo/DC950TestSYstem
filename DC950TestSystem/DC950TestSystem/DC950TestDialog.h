// DC950TestDialog.h : header file
//

#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "ReadOnlyEdit.h"
#include "PasswordEntryDlg.h"
#include "AdminDlg.h"


class CTestDialog : public CDialog
{
// Construction
public:
	CTestDialog(CWnd* pParent = NULL);	// standard constructor
	~CTestDialog();						// destructor

// Dialog Data
	enum { IDD = IDD_SERIALCTRLDEMO_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();	
	afx_msg void OnAdd();
	afx_msg void OnBack();
	afx_msg void OnText();
	
	HICON m_hIcon;	
	DECLARE_MESSAGE_MAP()
public:
	CStatic m_static_BarcodeLabel;
	CStatic m_static_TestResult;
	CStatic m_static_ComOut;
	CStatic m_static_ComIn;
	CStatic m_static_Line1;
	CStatic m_static_Line2;
	CStatic m_static_Line3;
	CStatic m_static_Line4;
	CStatic m_static_Line5;	
	CStatic m_static_Line6;
	CStatic m_static_DataIn;
	CStatic m_static_DataOut;
	CStatic m_static_TestName;
	CButton m_static_ButtonEnter;	
	CButton m_static_ButtonTest;
	CButton m_static_ButtonPass;
	CButton m_static_ButtonFail;
	CButton m_static_ButtonHalt;
	CButton m_static_ButtonPrevious;
	CButton m_static_optStandard;
	CButton m_static_optFilter;
	CButton m_static_modelGroup;
	CButton m_static_ButtonAdmin;
	CEdit   m_BarcodeEditBox;
	CEdit   m_EditBox_Log;
	
	PasswordEntry m_PED;
	Admin m_Admin;
	CEdit*  pEdit;

	CReadOnlyEdit	m_EditBox_Test1;
	CReadOnlyEdit	m_EditBox_Test2;
	CReadOnlyEdit	m_EditBox_Test3;
	CReadOnlyEdit	m_EditBox_Test4;
	CReadOnlyEdit	m_EditBox_Test5;
	CReadOnlyEdit	m_EditBox_Test6;
	CReadOnlyEdit	m_EditBox_Test7;
	CReadOnlyEdit	m_EditBox_Test8;
	CReadOnlyEdit	m_EditBox_Test9;
	CReadOnlyEdit	m_EditBox_Test10;
	#define NUM_TEST_EDIT_BOXES 10
		
	CStatusBar m_StatusBar;	
	UINT  stepNumber;
	UINT  subStepNumber;
	BOOL runFilterActuatorTest;
	
	HANDLE m_timerHandle;
	void timerHandler();
	void StartTimer();
	void StopTimer();				
	void CTestDialog::ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold);
	void CreateStatusBars();	
	void Testhandler();		
	void CTestDialog::SetDialogColor(COLORREF rgb);
	BOOL CTestDialog::PreTranslateMessage(MSG* pMsg);
	BOOL enableAdmin;
public:		
	afx_msg void OnBtnClickedTest();
	afx_msg void OnBtnClickedHalt();
	afx_msg void OnBtnClickedPrevious();	
	afx_msg void OnBtnClickedPass();
	afx_msg void OnBtnClickedFail();
	afx_msg void OnBtnClickedEnter();
	afx_msg void OnBtnClickedRadioSTD();
	afx_msg void OnBtnClickedRadioFilter();
	afx_msg void OnStnClickedStaticLine4();
	afx_msg void OnBnClickedButtonAdmin();
protected:
};
