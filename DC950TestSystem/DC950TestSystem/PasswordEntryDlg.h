// PasswordEntryDlg.h
#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


class PasswordEntry : public CDialog
{
public:
	PasswordEntry(CWnd* pParent = NULL);   // standard constructor

	enum { IDD = IDD_DIALOG_PWENTRY };
	CString	m_strPW;	
	public:
	virtual int DoModal(CString* pstrPW);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
protected:
	afx_msg void OnClose();
	virtual void OnCancel();
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()
private:
	CString* m_pstrPW;
public:
	afx_msg void OnEnChangeEditPassword();
};