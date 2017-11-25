// AdminDlg.cpp : implementation file for Admin Settings dialog box

#include "stdafx.h"
#include "Definitions.h"
#include <windows.h>
#include <iostream>
#include <sstream>
#include <string.h>
#include "AdminDlg.h"
#include "TestApp.h"

extern TestApp MyTestApp;	

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Admin dialog

extern long portNumberInterfaceBoard, portNumberACpowerSupply, portNumberMultiMeter;
extern float AllowableLampVoltageError, AllowableVrefError,
MinClosedFilterColorTemp, MaxClosedFilterColorTemp,
MinClosedFilterWavelength, MaxClosedFilterWavelength,
MinClosedFilterIrradiance, MaxClosedFilterIrradiance,
MinClosedFilterFWHM, MaxClosedFilterFWHM,
MinClosedFilterAmplitude, MinOpenColorTemp,
MaxOpenColorTemp, MinOpenIrradiance,
MaxOpenIrradiance;

Admin::Admin(CWnd* pParent /*=NULL*/) : CDialog(Admin::IDD, pParent)
{
	
}


void Admin::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);	
	DDX_Control(pDX, IDC_SPIN_COM_INTERFACE, m_SpinCOMInterface);
	DDX_Control(pDX, IDC_EDIT_COM_INTERFACE, m_EditCOMInterface);	
	DDX_Control(pDX, IDC_SPIN_COM_MULTIMETER, m_SpinCOMMultiMeter);
	DDX_Control(pDX, IDC_EDIT_COM_MULTIMETER, m_EditCOMMultiMeter);	
	DDX_Control(pDX, IDC_SPIN_COM_POWER_SUPPLY, m_SpinCOMPowerSupply);
	DDX_Control(pDX, IDC_EDIT_COM_POWER_SUPPLY, m_EditCOMPowerSupply);	

	DDX_Control(pDX, IDC_SPIN_LAMP_VOLTAGE,	m_SpinAllowableLampVoltageError);
	DDX_Control(pDX, IDC_SPIN_REF_VOLTAGE, 	m_SpinAllowableVrefError);
	DDX_Control(pDX, IDC_SPIN_MIN_COLOR_TEMP_FILTER, m_SpinMinClosedFilterColorTemp);
	DDX_Control(pDX, IDC_SPIN_MAX_COLOR_TEMP_FILTER, m_SpinMaxClosedFilterColorTemp);
	DDX_Control(pDX, IDC_SPIN_MIN_WAVELENGTH_FILTER, m_SpinMinClosedFilterWavelength);
	DDX_Control(pDX, IDC_SPIN_MAX_WAVELENGTH_FILTER, m_SpinMaxClosedFilterWavelength);
	DDX_Control(pDX, IDC_SPIN_MIN_IRRADIANCE_FILTER, m_SpinMinClosedFilterIrradiance);
	DDX_Control(pDX, IDC_SPIN_MAX_IRRADIANCE_FILTER, m_SpinMaxClosedFilterIrradiance);
	DDX_Control(pDX, IDC_SPIN_MIN_FWHM_FILTER, m_SpinMinClosedFilterFWHM);
	DDX_Control(pDX, IDC_SPIN_MAX_FWHM_FILTER, m_SpinMaxClosedFilterFWHM);

	DDX_Control(pDX, IDC_SPIN_MIN_COLOR_TEMP_OPEN, m_SpinMinOpenColorTemp);
	DDX_Control(pDX, IDC_SPIN_MAX_COLOR_TEMP_OPEN, m_SpinMaxOpenColorTemp);
	DDX_Control(pDX, IDC_SPIN_MIN_IRRADIANCE_OPEN, m_SpinMinOpenIrradiance);
	DDX_Control(pDX, IDC_SPIN_MAX_IRRADIANCE_OPEN, m_SpinMaxOpenIrradiance);

	DDX_Control(pDX, IDC_EDIT_MIN_COLOR_TEMP_OPEN, m_EditMinOpenColorTemp);
	DDX_Control(pDX, IDC_EDIT_MAX_COLOR_TEMP_OPEN, m_EditMaxOpenColorTemp);
	DDX_Control(pDX, IDC_EDIT_MIN_IRRADIANCE_OPEN, m_EditMinOpenIrradiance);
	DDX_Control(pDX, IDC_EDIT_MAX_IRRADIANCE_OPEN, m_EditMaxOpenIrradiance);


	DDX_Control(pDX, IDC_EDIT_LAMP_VOLTAGE,	m_EditAllowableLampVoltageError);
	DDX_Control(pDX, IDC_EDIT_REF_VOLTAGE, 	m_EditAllowableVrefError);
	DDX_Control(pDX, IDC_EDIT_MIN_COLOR_TEMP_FILTER, m_EditMinClosedFilterColorTemp);
	DDX_Control(pDX, IDC_EDIT_MAX_COLOR_TEMP_FILTER, m_EditMaxClosedFilterColorTemp);
	DDX_Control(pDX, IDC_EDIT_MIN_WAVELENGTH_FILTER, m_EditMinClosedFilterWavelength);
	DDX_Control(pDX, IDC_EDIT_MAX_WAVELENGTH_FILTER, m_EditMaxClosedFilterWavelength);
	DDX_Control(pDX, IDC_EDIT_MIN_IRRADIANCE_FILTER, m_EditMinClosedFilterIrradiance);
	DDX_Control(pDX, IDC_EDIT_MAX_IRRADIANCE_FILTER, m_EditMaxClosedFilterIrradiance);
	DDX_Control(pDX, IDC_EDIT_MIN_FWHM_FILTER, m_EditMinClosedFilterFWHM);
	DDX_Control(pDX, IDC_EDIT_MAX_FWHM_FILTER, m_EditMaxClosedFilterFWHM);

}
	

BEGIN_MESSAGE_MAP(Admin, CDialog)
	ON_WM_CLOSE()		
	ON_BN_CLICKED(IDOK, &Admin::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Admin::OnBnClickedCancel)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Admin message handlers

//int Admin::DoModal() 
//{
	// TODO: Add your specialized code here and/or call the base class
//	return (int) CDialog::DoModal();
//}

void Admin::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnClose();
}

/*
void Admin::OnCancel() 
{
	// TODO: Add extra cleanup here
	CDialog::OnCancel();
}

void Admin::OnOK() 
{
	// TODO: Add extra validation here
	// UpdateData(true);
	
	CDialog::OnOK();
}
*/


BOOL Admin::OnInitDialog() 
{
	CDialog::OnInitDialog();
	// UpdateData(false);	

	
	// TODO: Add extra initialization here
	// Initilize Numerical Spin Controls	
	m_SpinCOMInterface.SetBuddy(&m_EditCOMInterface);
	m_SpinCOMInterface.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMInterface.SetPos((float) portNumberInterfaceBoard);
	m_SpinCOMInterface.SetDelta(1);	
	
	m_SpinCOMMultiMeter.SetBuddy(&m_EditCOMMultiMeter);
	m_SpinCOMMultiMeter.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMMultiMeter.SetPos((float) portNumberMultiMeter);
	m_SpinCOMMultiMeter.SetDelta(1);		

	m_SpinCOMPowerSupply.SetBuddy(&m_EditCOMPowerSupply);
	m_SpinCOMPowerSupply.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMPowerSupply.SetPos((float) portNumberACpowerSupply);
	m_SpinCOMPowerSupply.SetDelta(1);		

	m_SpinAllowableLampVoltageError.SetBuddy(&m_EditAllowableLampVoltageError);
	m_SpinAllowableLampVoltageError.SetRange((float) 0.001, (float) 10.000);
	m_SpinAllowableLampVoltageError.SetPos(AllowableLampVoltageError);
	m_SpinAllowableLampVoltageError.SetDelta((float) 0.01);

	m_SpinAllowableVrefError.SetBuddy(&m_EditAllowableVrefError);
	m_SpinAllowableVrefError.SetRange((float) 0.001, (float) 1.000);
	m_SpinAllowableVrefError.SetPos(AllowableVrefError);
	m_SpinAllowableVrefError.SetDelta((float) 0.01);

	m_SpinMinClosedFilterColorTemp.SetBuddy(&m_EditMinClosedFilterColorTemp);
	m_SpinMinClosedFilterColorTemp.SetRange(1, 1000);
	m_SpinMinClosedFilterColorTemp.SetPos(MinClosedFilterColorTemp);
	m_SpinMinClosedFilterColorTemp.SetDelta(1);

	m_SpinMaxClosedFilterColorTemp.SetBuddy(&m_EditMaxClosedFilterColorTemp);
	m_SpinMaxClosedFilterColorTemp.SetRange(1, 1000);
	m_SpinMaxClosedFilterColorTemp.SetPos(MaxClosedFilterColorTemp);
	m_SpinMaxClosedFilterColorTemp.SetDelta(1);

	m_SpinMinClosedFilterWavelength.SetBuddy(&m_EditMinClosedFilterWavelength);
	m_SpinMinClosedFilterWavelength.SetRange(1, 1000);
	m_SpinMinClosedFilterWavelength.SetPos(MinClosedFilterWavelength); 
	m_SpinMinClosedFilterWavelength.SetDelta(1);

	m_SpinMaxClosedFilterWavelength.SetBuddy(&m_EditMaxClosedFilterWavelength);
	m_SpinMaxClosedFilterWavelength.SetRange(1, 1000);
	m_SpinMaxClosedFilterWavelength.SetPos(MaxClosedFilterWavelength); 
	m_SpinMaxClosedFilterWavelength.SetDelta(1);

	m_SpinMinClosedFilterIrradiance.SetBuddy(&m_EditMinClosedFilterIrradiance);
	m_SpinMinClosedFilterIrradiance.SetRange(1, 1000);
	m_SpinMinClosedFilterIrradiance.SetPos(MinClosedFilterIrradiance); 
	m_SpinMinClosedFilterIrradiance.SetDelta(1);

	m_SpinMaxClosedFilterIrradiance.SetBuddy(&m_EditMaxClosedFilterIrradiance);
	m_SpinMaxClosedFilterIrradiance.SetRange(1, 1000);
	m_SpinMaxClosedFilterIrradiance.SetPos(MaxClosedFilterIrradiance); 
	m_SpinMaxClosedFilterIrradiance.SetDelta(1);

	m_SpinMinClosedFilterFWHM.SetBuddy(&m_EditMinClosedFilterFWHM);
	m_SpinMinClosedFilterFWHM.SetRange(1, 1000);
	m_SpinMinClosedFilterFWHM.SetPos(MinClosedFilterFWHM); 
	m_SpinMinClosedFilterFWHM.SetDelta(1);

	m_SpinMaxClosedFilterFWHM.SetBuddy(&m_EditMaxClosedFilterFWHM);
	m_SpinMaxClosedFilterFWHM.SetRange(1, 1000);
	m_SpinMaxClosedFilterFWHM.SetPos(MaxClosedFilterFWHM); 
	m_SpinMaxClosedFilterFWHM.SetDelta(1);

	m_SpinMinOpenIrradiance.SetBuddy(&m_EditMinOpenIrradiance);
	m_SpinMinOpenIrradiance.SetRange(1, 1000);
	m_SpinMinOpenIrradiance.SetPos(MinOpenIrradiance); 
	m_SpinMinOpenIrradiance.SetDelta(1);

	m_SpinMaxOpenIrradiance.SetBuddy(&m_EditMaxOpenIrradiance);
	m_SpinMaxOpenIrradiance.SetRange(1, 1000);
	m_SpinMaxOpenIrradiance.SetPos(MaxOpenIrradiance); 
	m_SpinMaxOpenIrradiance.SetDelta(1);

	m_SpinMinOpenColorTemp.SetBuddy(&m_EditMinOpenColorTemp);
	m_SpinMinOpenColorTemp.SetRange(1, 1000);
	m_SpinMinOpenColorTemp.SetPos(MinOpenColorTemp); 
	m_SpinMinOpenColorTemp.SetDelta(1);

	m_SpinMaxOpenColorTemp.SetBuddy(&m_EditMaxOpenColorTemp);
	m_SpinMaxOpenColorTemp.SetRange(1, 1000);
	m_SpinMaxOpenColorTemp.SetPos(MaxOpenColorTemp); 
	m_SpinMaxOpenColorTemp.SetDelta(1);
	

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


void Admin::OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult)
{
	// This feature requires Windows XP or greater.
	// The symbol _WIN32_WINNT must be >= 0x0501.
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void Admin::OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void Admin::OnPaint() 
{
	
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0); 

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		// dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
	
}

void Admin::OnBnClickedOk()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to save new settings", (LPCTSTR)"SAVE NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) {	
		portNumberInterfaceBoard = (long) m_SpinCOMInterface.GetPos();
		portNumberMultiMeter = (long) m_SpinCOMMultiMeter.GetPos();
		portNumberACpowerSupply = (long) m_SpinCOMPowerSupply.GetPos();
	
		AllowableLampVoltageError = m_SpinAllowableLampVoltageError.GetPos();
		AllowableVrefError = m_SpinAllowableVrefError.GetPos();
		MinClosedFilterColorTemp = m_SpinMinClosedFilterColorTemp.GetPos();
		MaxClosedFilterColorTemp = m_SpinMaxClosedFilterColorTemp.GetPos();
		MinClosedFilterWavelength = m_SpinMinClosedFilterWavelength.GetPos();
		MaxClosedFilterWavelength = m_SpinMaxClosedFilterWavelength.GetPos();
		MinClosedFilterIrradiance = m_SpinMinClosedFilterIrradiance.GetPos();
		MaxClosedFilterIrradiance = m_SpinMaxClosedFilterIrradiance.GetPos();
		MinClosedFilterFWHM = m_SpinMinClosedFilterFWHM.GetPos();
		MaxClosedFilterFWHM = m_SpinMaxClosedFilterFWHM.GetPos(); 
		MinOpenColorTemp = m_SpinMinOpenColorTemp.GetPos(); 
		MaxOpenColorTemp = m_SpinMaxOpenColorTemp.GetPos(); 
		MinOpenIrradiance = m_SpinMinOpenIrradiance.GetPos();
		MaxOpenIrradiance = m_SpinMaxOpenIrradiance.GetPos();
		MyTestApp.write_INI_File();
		CDialog::OnOK();
	}
}

void Admin::OnBnClickedCancel()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to quit without saving new settings", (LPCTSTR)"DISGARD NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) CDialog::OnCancel();
}

