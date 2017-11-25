//#if !defined(AFX_PASSWORDENTRY_DLG_H__C2118394_6C77_4314_9CA1_C8495031C9D0__INCLUDED_)
//#define AFX_PASSWORDENTRY_DLG_H__C2118394_6C77_4314_9CA1_C8495031C9D0__INCLUDED_

#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
/////////////////////////////////////////////////////////////////////////////
// PasswordEntry_Dlg dialog



class PasswordEntry_Dlg : public CDialog
{
// Construction
public:
	PasswordEntry_Dlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(PasswordEntry_Dlg)
	enum { IDD = IDD_DIALOG_PWENTRY };
	CString	m_strPW;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(PasswordEntry_Dlg)
	public:
	virtual int DoModal(CString* pstrPW);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(PasswordEntry_Dlg)
	afx_msg void OnClose();
	virtual void OnCancel();
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	CString* m_pstrPW;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

//#endif // !defined(AFX_PASSWORDENTRY_DLG_H__C2118394_6C77_4314_9CA1_C8495031C9D0__INCLUDED_)
