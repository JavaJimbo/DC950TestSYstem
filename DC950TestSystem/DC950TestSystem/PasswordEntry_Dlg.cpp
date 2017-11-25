// PasswordEntry_Dlg.cpp : implementation file
//

#include "stdafx.h"
// #include "Calibrator_TestAndCal.h"
#include "SerialCtrlDemo.h"
#include "SerialCtrlDemoDlg.h"
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

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// PasswordEntry_Dlg dialog


PasswordEntry_Dlg::PasswordEntry_Dlg(CWnd* pParent /*=NULL*/) : CDialog(PasswordEntry_Dlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(PasswordEntry_Dlg)
	m_strPW = _T("");
	//}}AFX_DATA_INIT
}


void PasswordEntry_Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(PasswordEntry_Dlg)
	DDX_Text(pDX, IDC_EDIT_PASSWORD, m_strPW);
	DDV_MaxChars(pDX, m_strPW, 32);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(PasswordEntry_Dlg, CDialog)
	//{{AFX_MSG_MAP(PasswordEntry_Dlg)
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// PasswordEntry_Dlg message handlers

int PasswordEntry_Dlg::DoModal(CString* pstrPW) 
{
	// TODO: Add your specialized code here and/or call the base class
	m_pstrPW = pstrPW;
	return CDialog::DoModal();
}

void PasswordEntry_Dlg::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	m_pstrPW = NULL;	
	CDialog::OnClose();
}

void PasswordEntry_Dlg::OnCancel() 
{
	// TODO: Add extra cleanup here
	m_strPW = "NULL";
	CDialog::OnCancel();
}

void PasswordEntry_Dlg::OnOK() 
{
	// TODO: Add extra validation here
	UpdateData(true);
	*m_pstrPW = m_strPW;
	CDialog::OnOK();
}

BOOL PasswordEntry_Dlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	m_strPW = "";
	UpdateData(false);
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
