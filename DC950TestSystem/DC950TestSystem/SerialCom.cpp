/* SerialCom.cpp - low level routines for sending and receiving data on serial ports, as well as opening and closing.
 * Written in Visual C++ for DC950
 *
 */

// NOTE: INCUDES MUST BE IN THIS ORDER!!!
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"

extern UINT16  CRCcalculate(char *ptrPacket, BOOL addCRCtoPacket);
extern BOOL CRCcheck(char *ptrPacket);
extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

BOOL TestApp::openSerialPort (const char *ptrPortName, HANDLE *ptrPortHandle) {
			DCB serialPortConfig;
			BOOL tryAgain = FALSE;		

			if (ptrPortName == NULL) return FALSE;
			do {
					*ptrPortHandle = CreateFile((LPCSTR)ptrPortName,  // Specify port device: default "COM1"		
					GENERIC_READ | GENERIC_WRITE,       // Specify mode that open device.
					0,                                  // the devide isn't shared.
					NULL,                               // the object gets a default security.
					OPEN_EXISTING,                      // Specify which action to take on file. 
					0,                                  // default.
					NULL);                              // default.

					char errorMessage[BUFFERSIZE];
					sprintf_s(errorMessage, "ERROR: cannot open %s", ptrPortName);

				if (GetCommState(*ptrPortHandle, &serialPortConfig) == 0) {
					tryAgain = DisplayMessageBox(errorMessage, "Check USB connections and try again.", 2);
					if (!tryAgain) return FALSE;
				}
				else break;   // SUCCESS! Port opened OK, no need to try opening it again
				msDelay(100);
			} while (tryAgain);

			DCB dcb;
			dcb.BaudRate = CBR_9600;			// Fix baud rate at 9600, 1 stop bit, no parity, 8 data bits
			dcb.StopBits = ONESTOPBIT;
			dcb.Parity = NOPARITY;
			dcb.ByteSize = DATABITS_8;

			// Assign user parameter.
			serialPortConfig.BaudRate = dcb.BaudRate;    // Specify buad rate of communicaiton.
			serialPortConfig.StopBits = dcb.StopBits;    // Specify stopbit of communication.
			serialPortConfig.Parity = dcb.Parity;        // Specify parity of communication.
			serialPortConfig.ByteSize = dcb.ByteSize;    // Specify  byte of size of communication.

														 // Set current configuration of serial communication port.		
														 // if (SetCommState(m_NewPortHandle, &serialPortConfig) == 0)
			if (SetCommState(*ptrPortHandle, &serialPortConfig) == 0)
			{
				AfxMessageBox((LPCTSTR)"PROGRAM ERROR: Set configuration port has problem.");
				return FALSE;
			}

			// instance an object of COMMTIMEOUTS.
			COMMTIMEOUTS comTimeOut;
			comTimeOut.ReadIntervalTimeout = 3;
			comTimeOut.ReadTotalTimeoutMultiplier = 3;
			comTimeOut.ReadTotalTimeoutConstant = 2;
			comTimeOut.WriteTotalTimeoutMultiplier = 3;
			comTimeOut.WriteTotalTimeoutConstant = 2;

			SetCommTimeouts(*ptrPortHandle, &comTimeOut);		// set the time-out parameter into device control.												
			return TRUE;  // Return TRUE to indicate port is successfully opened and configured
	}

	BOOL TestApp::closeSerialPort(HANDLE ptrPortHandle) {
		if (ptrPortHandle == NULL) 
			return TRUE;

		if (!CloseHandle(ptrPortHandle))
		{
			AfxMessageBox((LPCTSTR)"Port Closing failed.");
			return FALSE;
		}
		ptrPortHandle = NULL;
		return(TRUE);
	}


	
	
	BOOL TestApp::WriteSerialPort (HANDLE ptrPortHandle, char *ptrPacket) {
		int length;
		int trial = 0;
		int numBytesWritten = 0;		

		if (ptrPacket == NULL || ptrPortHandle == NULL) return (FALSE);
		
		length = (int) strlen(ptrPacket);					
		if (length >= BUFFERSIZE) {
			// intError = SYSTEM_ERROR;
			return (FALSE);
		}		

		do {
			trial++;
			if (WriteFile(ptrPortHandle, ptrPacket, length, (LPDWORD) &numBytesWritten, NULL)) break;
			msDelay(100);
		} while (trial < MAXTRIES && numBytesWritten == 0);

		if (trial >= MAXTRIES) return FALSE;
		else return TRUE;
	}
	
		
	BOOL TestApp::ReadSerialPort(HANDLE ptrPortHandle, char *ptrPacket) {
		int totalBytesRead = 0;
		DWORD numBytesRead;
		char inBytes[BUFFERSIZE];
		ptrPacket[0] = '\0';
		int trial = 0;

		if (ptrPacket == NULL) {			
			return (FALSE);
		}

		trial = 0;
		ptrPacket[0] = '\0';
		do {
			trial++;		
			msDelay(100);
			if (ReadFile(ptrPortHandle, inBytes, BUFFERSIZE, &numBytesRead, NULL)) {
				if (numBytesRead > 0 && numBytesRead < BUFFERSIZE) {
					inBytes[numBytesRead] = '\0';
					strcat_s(ptrPacket, BUFFERSIZE, inBytes);
					if (strchr(inBytes, '\r')) break;	
					if (strchr(inBytes, '\n')) break;
				}
			}
			
		} while (trial < MAXTRIES);

		if (trial >= MAXTRIES) return (FALSE);		
		else return (TRUE);		
	}
	
	BOOL TestApp::sendReceiveSerial(int COMdevice,  char *outPacket, char *inPacket, BOOL expectReply)
	{
		HANDLE ptrPortHandle = NULL;

		switch (COMdevice) {
		case HP_METER:
			ptrPortHandle = handleHPmultiMeter;
			break;
		case AC_POWER_SUPPLY:
			ptrPortHandle = handleACpowerSupply;
			break;
		case INTERFACE_BOARD:
			ptrPortHandle = handleInterfaceBoard;
		default:
			break;
		}		

		if (ptrPortHandle == NULL || outPacket == NULL) {
			return (FALSE);
		}

		DisplaySerialComData(DATAOUT, outPacket);
		if (COMdevice == INTERFACE_BOARD) 
			CRCcalculate(outPacket, TRUE);

		if (outPacket == NULL) return (FALSE);

		DisplaySerialComData(DATAIN, "");

		if (!WriteSerialPort(ptrPortHandle, outPacket)) {
			DisplaySerialComData(DATAOUT, "COM PORT ERROR");
			return (FALSE);
		}

		if (!expectReply) return TRUE;

		if (inPacket == NULL) {
			char temp[BUFFERSIZE];
			inPacket = temp;
		}

		if (!ReadSerialPort(ptrPortHandle, inPacket)) {
			DisplaySerialComData(DATAIN, "COM PORT ERROR");
			return (FALSE);
		}

		if (COMdevice == INTERFACE_BOARD) {
			if (!CRCcheck(inPacket)) {
				DisplaySerialComData(DATAIN, "DATA IN CRC ERROR");
				return (FALSE);
			}			
		}
		DisplaySerialComData(DATAIN, inPacket);
		return (TRUE);		
	}


	