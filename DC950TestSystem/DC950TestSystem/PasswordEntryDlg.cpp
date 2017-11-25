// PasswordEntryDlg.cpp : implementation file for password dialog box
// This is used to access the Admin dialog box.


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
#include "PasswordEntryDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// PasswordEntry dialog


PasswordEntry::PasswordEntry(CWnd* pParent /*=NULL*/) : CDialog(PasswordEntry::IDD, pParent)
{
	m_strPW = _T("");
}


void PasswordEntry::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_PASSWORD, m_strPW);
	DDV_MaxChars(pDX, m_strPW, 32);
}


BEGIN_MESSAGE_MAP(PasswordEntry, CDialog)
	ON_WM_CLOSE()
	ON_EN_CHANGE(IDC_EDIT_PASSWORD, &PasswordEntry::OnEnChangeEditPassword)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// PasswordEntry message handlers

int PasswordEntry::DoModal(CString* pstrPW) 
{
	// TODO: Add your specialized code here and/or call the base class
	m_pstrPW = pstrPW;
	return (int) CDialog::DoModal();
}

void PasswordEntry::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	m_pstrPW = NULL;	
	CDialog::OnClose();
}

void PasswordEntry::OnCancel() 
{
	// TODO: Add extra cleanup here
	m_strPW = "NULL";
	CDialog::OnCancel();
}

void PasswordEntry::OnOK() 
{
	// TODO: Add extra validation here
	UpdateData(true);
	*m_pstrPW = m_strPW;
	CDialog::OnOK();	
}

BOOL PasswordEntry::OnInitDialog() 
{
	CDialog::OnInitDialog();
	m_strPW = "";
	UpdateData(false);
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


void PasswordEntry::OnEnChangeEditPassword()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
}
