

// AVSPectrometer - routines for communicating with Avantes spectrometer
// Inlcudes function for receiving and sending configuration information
// as well as initiating scans amd fetching readings.

#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"
#include "avaspec.h"
#include <fstream>
#include <iostream>

using namespace std;

#define MAXPIXELS 2048

int m_NrDevices = 0;
long handleSpectrometer = INVALID_AVS_HANDLE_VALUE;
unsigned int numberOfPixels = 0;
double Lambda[MAXPIXELS];
double Spectrum[MAXPIXELS];
uint16	m_StartPixel = 0;
uint16	m_StopPixel = 0;
uint16	m_NumberOfScans = -1;
float fltIntegrationTime = 100.0;
unsigned long numberOfAverages = 30; // 300;
char* l_pData;
AvsIdentityType*    spectrometerID;

void clearDataArrays() {
	for (int i = 0; i < MAXPIXELS; i++) {
		Lambda[i] = 0.0;
		Spectrum[i] = 0.0;
	}
}

BOOL TestApp::InitializeSpectrometer(CTestDialog *ptrDialog) {	
	long l_Port = 0;	
	unsigned int        l_Size = 0;
	unsigned int        l_RequiredSize = 0;

#ifdef SIMULMODE
	return TRUE;
#endif

	msDelay(100);
	l_Port = AVS_Init(0);  	
	
	if (l_Port > 0)
	{
		// UpdateList();
		msDelay(100);
		m_NrDevices = AVS_UpdateUSBDevices();		
		if (m_NrDevices != 1) return FALSE;
		msDelay(100);

		m_NrDevices = AVS_GetNrOfDevices();
		l_RequiredSize = m_NrDevices * sizeof(AvsIdentityType);

		if (l_RequiredSize > 0)
		{
			delete[] l_pData;
			l_pData = new char[l_RequiredSize];
			l_Size = l_RequiredSize;
			spectrometerID = (AvsIdentityType*)l_pData;
			m_NrDevices = AVS_GetList(l_Size, &l_RequiredSize, spectrometerID);
		}
		if (m_NrDevices != 1) return FALSE;
		else return TRUE;		
	}
	else AVS_Done();		
	
	return FALSE;
}

BOOL TestApp::ActivateSpectrometer(CTestDialog *ptrDialog) {

#ifdef SIMULMODE 
	return TRUE;
#endif
	
	long newHandle = AVS_Activate(spectrometerID);

	if (newHandle == INVALID_AVS_HANDLE_VALUE) return FALSE;
	else handleSpectrometer = newHandle;

	if (getSpectrometerConfiguration(ptrDialog)) return TRUE;
	else return FALSE;
}


BOOL TestApp::getSpectrometerConfiguration(CTestDialog *ptrDialog) {
	DeviceConfigType* l_pDeviceData = NULL;
	unsigned short intNumberOfPixels = 0;
	unsigned int l_Size;

#ifdef SIMULMODE
	return TRUE;
#endif
	if (handleSpectrometer == INVALID_AVS_HANDLE_VALUE) return FALSE;

	char a_Fpga[16];
	char a_As5216[16];
	char a_Dll[16];

	int l_Res = AVS_GetVersionInfo(handleSpectrometer, a_Fpga, a_As5216, a_Dll);

	if (ERR_SUCCESS == l_Res)
	{		
		DisplayText(ptrDialog, 1, a_Fpga);
		DisplayText(ptrDialog, 1, a_As5216);
		DisplayText(ptrDialog, 1, a_Dll);
	}

	if (ERR_SUCCESS == AVS_GetNumPixels(handleSpectrometer, &intNumberOfPixels))
			numberOfPixels = (unsigned int) intNumberOfPixels;

	l_Res = AVS_GetParameter(handleSpectrometer, 0, &l_Size, l_pDeviceData);

	if (l_Res == ERR_INVALID_SIZE)
	{
		l_pDeviceData = (DeviceConfigType*)new char[l_Size];
	}

	l_Res = AVS_GetParameter(handleSpectrometer, l_Size, &l_Size, l_pDeviceData);

	if (l_Res != ERR_SUCCESS) return FALSE;

	m_StartPixel = l_pDeviceData->m_StandAlone.m_Meas.m_StartPixel;
	m_StopPixel = l_pDeviceData->m_StandAlone.m_Meas.m_StopPixel;
	if (m_StopPixel > MAXPIXELS) m_StopPixel = MAXPIXELS;
	
	clearDataArrays();

	if (ERR_SUCCESS == AVS_GetLambda(handleSpectrometer, Lambda)) return TRUE;
	else return FALSE;
}


BOOL TestApp::startMeasurement()
{
	double l_NanoSec = 0.0;
	long l_Ticks = 0;
	long l_Dif = 0;
	unsigned int l_Time = 0;
	uint32 l_FPGAClkCycles = 0;
	long l_Res = 0;
	unsigned char l_Status = 0;
	MeasConfigType l_PrepareMeasData;

#ifdef SIMULMODE
	return TRUE;
#endif		
		if (spectrometerID->Status == USB_IN_USE_BY_APPLICATION)
		{
			// StartPixel				
			l_PrepareMeasData.m_StartPixel = m_StartPixel;
			// StopPixel
			l_PrepareMeasData.m_StopPixel = m_StopPixel;
			// IntegrationTime			
			l_PrepareMeasData.m_IntegrationTime = fltIntegrationTime;
			// IntegrationDelay
			l_NanoSec = -20.83;
			l_FPGAClkCycles = (uint32)(6.0 * (l_NanoSec + 20.84) / 125.0);
			l_PrepareMeasData.m_IntegrationDelay = l_FPGAClkCycles;
			// Number of Averages
			l_PrepareMeasData.m_NrAverages = numberOfAverages;
			// DarkCorrectionType
			l_PrepareMeasData.m_CorDynDark.m_Enable = 1;
			l_PrepareMeasData.m_CorDynDark.m_ForgetPercentage = 100;
			// SmoothingType
			l_PrepareMeasData.m_Smoothing.m_SmoothPix = 1;
			l_PrepareMeasData.m_Smoothing.m_SmoothModel = 0;
			// SaturationDetection
			l_PrepareMeasData.m_SaturationDetection = 1;
			// TriggerType
			l_PrepareMeasData.m_Trigger.m_Mode = 0;
			l_PrepareMeasData.m_Trigger.m_Source = 0;
			l_PrepareMeasData.m_Trigger.m_Source = 1;
			l_PrepareMeasData.m_Trigger.m_SourceType = 0;
			l_PrepareMeasData.m_Trigger.m_SourceType = 1;

			// ControlSettingsType
			l_PrepareMeasData.m_Control.m_StrobeControl = 0;
			l_NanoSec = 0;
			l_FPGAClkCycles = (uint32)(6.0 * l_NanoSec / 125.0);
			l_PrepareMeasData.m_Control.m_LaserDelay = l_FPGAClkCycles;
			l_NanoSec = 0;
			l_FPGAClkCycles = (uint32)(6.0 * l_NanoSec / 125.0);
			l_PrepareMeasData.m_Control.m_LaserWidth = l_FPGAClkCycles;
			l_PrepareMeasData.m_Control.m_LaserWaveLength = 0;
			l_PrepareMeasData.m_Control.m_StoreToRam = 0;

			if (ERR_SUCCESS != AVS_PrepareMeasure(handleSpectrometer, &l_PrepareMeasData))
				return FALSE;			

			// Start Measurement
			m_StartPixel = l_PrepareMeasData.m_StartPixel;
			m_StopPixel = l_PrepareMeasData.m_StopPixel;
			
			clearDataArrays();
			if (ERR_SUCCESS != AVS_GetLambda(handleSpectrometer, Lambda))
				return FALSE;
			if (ERR_SUCCESS != AVS_Measure(handleSpectrometer, NULL, m_NumberOfScans))
				return FALSE;
			else return TRUE;
		}
		else return FALSE;
}


BOOL TestApp::stopMeasurement() {
#ifdef SIMULMODE
	return TRUE;
#endif

	if (spectrometerID->Status == USB_IN_USE_BY_APPLICATION)
	{
		if (ERR_SUCCESS != AVS_StopMeasure(handleSpectrometer))
			return FALSE;
		else return TRUE;		
	}
	else return FALSE;
}

BOOL TestApp::spectrometerDataReady() {
#ifdef SIMULMODE
	return TRUE;
#endif

	if (AVS_PollScan(handleSpectrometer))
		return TRUE;
	else return FALSE;
}

BOOL TestApp::CloseSpectrometer()
{
#ifdef SIMULMODE
	return TRUE;
#endif

	if (handleSpectrometer) AVS_Done();
	return TRUE;
}




BOOL TestApp::getWavelengthAndIrradiance(float *peakIrradiance, float *wavelength)
{
#ifdef SIMULMODE
	return TRUE;
#endif

	uint16 i;
	*peakIrradiance = 0;
	*wavelength = 0;
	float spectrumValue = 0;
	
	if (m_StopPixel >= MAXPIXELS) return FALSE;

	for (i = 0; i < m_StopPixel; i++) {
		spectrumValue = (float) Spectrum[i];
		if (*peakIrradiance < spectrumValue)
		{
			*peakIrradiance = spectrumValue;
			*wavelength = (float) Lambda[i];
		}
	}	
	return TRUE;
}

BOOL TestApp::getMinAmplitude(float *minAmplitude)
{
#ifdef SIMULMODE
	return TRUE;
#endif

	uint16 i;
	float spectrumValue = 0;

	// Initialize minimum spectrum reading with arbitrarily high value:
	*minAmplitude = (float) 0xFFFF;

	if (m_StopPixel >= MAXPIXELS) return FALSE;

	for (i = 0; i < m_StopPixel; i++) {
		spectrumValue = (float)Spectrum[i];
		if (*minAmplitude > spectrumValue)
			*minAmplitude = spectrumValue;
	}	
	return TRUE;
}


BOOL TestApp::writeSpectrometerDataToFile(const char *filename) 
{
#ifdef SIMULMODE
	return TRUE;
#endif

	ofstream myfile;
	uint16 i;
	char strSpectrometerData[MAXSTRING];
	float spectrumValue, lambdaValue;
		
	if (m_StopPixel >= MAXPIXELS) return FALSE;

	myfile.open(filename);
	for (i = 0; i < m_StopPixel; i++) {
		
		spectrumValue = (float) Spectrum[i];
		lambdaValue = (float) Lambda[i];
		sprintf_s (strSpectrometerData, MAXSTRING, ".2f, .2f\n", lambdaValue, spectrumValue);
		myfile << strSpectrometerData;
		// myfile << lambdaValue << "," << spectrumValue << "\n";
	}
	myfile.close();
	return TRUE;
}



BOOL TestApp::getSpectrometerData(CTestDialog *ptrDialog, float *wavelength, float *minAmplitude, float *peakIrradiance, float *FWHM, float *colorTemperature) {
#ifdef SIMULMODE
	return TRUE;
#endif

	uint32 l_Time = 0;

	if (ERR_SUCCESS == AVS_GetScopeData(handleSpectrometer, &l_Time, Spectrum)) {
		writeSpectrometerDataToFile("Spectrum.txt");

		if (!getWavelengthAndIrradiance(peakIrradiance, wavelength)) return FALSE;
		if (!getMinAmplitude(minAmplitude)) return FALSE;

		char strText[MAXSTRING];
		sprintf_s(strText, MAXSTRING, "Peak Irradiance: %.2f", *peakIrradiance);
		DisplayText(ptrDialog, 2, strText);

		sprintf_s(strText, MAXSTRING, "Wavelength: %.2f", *wavelength);
		DisplayText(ptrDialog, 3, strText);

		sprintf_s(strText, MAXSTRING, "Min amplitude: %.2f", *minAmplitude);
		DisplayText(ptrDialog, 4, strText);

		*FWHM = 14;
		sprintf_s(strText, MAXSTRING, "FWHM: %.2f", *FWHM);
		DisplayText(ptrDialog, 4, strText);

		*colorTemperature = 123;
		sprintf_s(strText, MAXSTRING, "ColorTemperature: %.2f", *FWHM);
		DisplayText(ptrDialog, 5, strText);

		return TRUE;
	}
	else return FALSE;
}
