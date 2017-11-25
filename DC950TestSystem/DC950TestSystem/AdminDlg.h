// PasswordEntry.h
#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "NumSpinCtrl.h"
#include "NumEdit.h"

class Admin : public CDialog
{
public:
	Admin(CWnd* pParent = NULL);   // standard constructor

	enum { IDD = IDD_DIALOG_ADMIN };
	CNumSpinCtrl	m_spinValue;

	public:
	// virtual int DoModal();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
protected:
	afx_msg void OnClose();
	afx_msg void OnPaint();
	// virtual void OnCancel();
	// virtual void OnOK();
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()
	CNumSpinCtrl m_SpinCOMInterface, m_SpinCOMMultiMeter, m_SpinCOMPowerSupply,
		m_SpinAllowableLampVoltageError, m_SpinAllowableVrefError,
		m_SpinMinClosedFilterColorTemp, m_SpinMaxClosedFilterColorTemp,
		m_SpinMinClosedFilterWavelength, m_SpinMaxClosedFilterWavelength,
		m_SpinMinClosedFilterIrradiance, m_SpinMaxClosedFilterIrradiance,
		m_SpinMinClosedFilterFWHM, m_SpinMaxClosedFilterFWHM,
		m_SpinMinClosedFilterAmplitude,
		m_SpinMinOpenColorTemp, m_SpinMaxOpenColorTemp,
		m_SpinMinOpenIrradiance, m_SpinMaxOpenIrradiance;


	CNumEdit m_EditCOMInterface, m_EditCOMMultiMeter, m_EditCOMPowerSupply,
		m_EditAllowableLampVoltageError, m_EditAllowableVrefError,
		m_EditMinClosedFilterColorTemp, m_EditMaxClosedFilterColorTemp,
		m_EditMinClosedFilterWavelength, m_EditMaxClosedFilterWavelength,
		m_EditMinClosedFilterIrradiance, m_EditMaxClosedFilterIrradiance,
		m_EditMinClosedFilterFWHM, m_EditMaxClosedFilterFWHM,
		m_EditMinClosedFilterAmplitude,
		m_EditMinOpenColorTemp, m_EditMaxOpenColorTemp,
		m_EditMinOpenIrradiance, m_EditMaxOpenIrradiance;

private:
public:
	afx_msg void OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
};