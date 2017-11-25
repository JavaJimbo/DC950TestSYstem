#pragma once

#ifndef DEFINITIONS_H
#define DEFINITIONS_H

#define MAXINIVALUES 32
#define MAX_PORTNUMBER 50

#define BLACK RGB(0,0,0)
#define YELLOW RGB(255,255,0)
#define RED RGB(255,0,0)
#define GREEN RGB(0,255,0)

#define MAXLOG 2048
#define SIMULMODE 
#define MAX_AC_VOLTAGE 254

#define FILTER_CLOSED 1
#define FILTER_OPEN 0

// #include "Definitions.h"
#define CHARWIDTH 340
#define MAXSTRING 128
#define MAX_PWM 6215
#define MULTIMETER 1
#define BUFFERSIZE 128
#define FONTHEIGHT 26
#define FONTWIDTH 10
#define MAXTRIES 5
#define MULTIMETER 1

#define TOTAL_STEPS 14
#define FINAL_FAIL TOTAL_STEPS-1
#define FINAL_PASS (TOTAL_STEPS-2)

#define STARTUP 0
#define LEFTPANEL 0
#define RIGHTPANEL 1


#define DATAIN 0
#define DATAOUT 1

/*
#define MIN_CLOSED_COLOR_TEMPERATURE (float) 100
#define MAX_CLOSED_COLOR_TEMPERATURE (float) 200
#define MIN_CLOSED_WAVELENGTH (float) 938
#define MAX_CLOSED_WAVELENGTH (float) 943
#define MIN_CLOSED_IRRADIANCE (float) 200 // 100
#define MAX_CLOSED_IRRADIANCE (float) 300 // 150
#define MIN_CLOSED_FWHM (float) 10
#define MAX_CLOSED_FWHM (float) 18

#define MIN_OPEN_COLOR_TEMPERATURE (float) 100
#define MAX_OPEN_COLOR_TEMPERATURE (float) 200
#define MIN_OPEN_WAVELENGTH (float) 938
#define MAX_OPEN_WAVELENGTH (float) 943
#define MIN_OPEN_IRRADIANCE (float) 200 // 100
#define MAX_OPEN_IRRADIANCE (float) 300 // 150
#define MIN_OPEN_FWHM (float) 10
#define MAX_OPEN_FWHM (float) 18
*/

enum DATA_COLUMNS {
	TEST_DATE = 0,
	SERIAL_NUMBER,	
	HI_POT,
	GROUND_BOND,
	LAMP_OFF,
	LAMP_ON,
	VREF_TEST,
	REM_VOLT_0,
	REM_VOLT_1,
	REM_VOLT_2,
	REM_VOLT_3,
	REM_VOLT_4,
	REM_VOLT_5,
	INHIBIT,		
	SWEEP_0,
	SWEEP_1,
	SWEEP_2,
	SWEEP_3,
	SWEEP_4,
	SWEEP_5,		
	TTL_OUT,	
	IRRADIANCE_CLOSED,
	WAVELENGTH_CLOSED,
	MIN_AMPLITUDE_CLOSED,
	FWHM_CLOSED,
	COLOR_TEMP_CLOSED,
	IRRADIANCE_OPEN,
	COLOR_TEMP_OPEN,	
	FINAL_RESULT,
	TOTAL_COLUMNS
};

#define END_STANDARD_UNIT_TESTS 9

#define SERIAL_NUMBER_COLUMN SERIAL_NUMBER
#define TEST_DATE_COLUMN TEST_DATE

#define FIRST_TEST_COLUMN SCAN_BARCODE
#define FIRST_FILTER_TEST_COLUMN ACTUATOR_OPEN

enum MULTIMETER_INPUTS {
	LAMP = 0,
	VREF,
	CONTROL_VOLTAGE
};

enum COM_PORTS {
	INTERFACE_BOARD = 0,
	HP_METER,
	AC_POWER_SUPPLY
};

enum ERROR_CODES {
	NO_ERRORS = 0,
	PORT_ERROR,
	TIMEOUT_ERROR,
	RESPONSE_ERROR,
	CRC_ERROR,
	SYSTEM_ERROR	
};

enum TIMER_STATES {
	TIMER_PAUSED = 0,
	TIMER_RUN
};

enum STATUS {
	NOT_DONE_YET,	
	PASS,
	FAIL
};

enum STEP_TYPE {			
	MANUAL = 0,
	AUTO
};

#define REMOTE_TEST 6
#define BARCODE_SCAN 1

struct TestStep
{
public:
	char *lineOne;
	char *lineTwo;
	char *lineThree;
	char *lineFour;
	char *lineFive;
	char *lineSix;
	char *testName;
	int  dataColNum;
	UINT Status;
	int  stepType;
	int  testID;
	BOOL enableENTER;
	BOOL enablePREVIOUS;
	BOOL enableHALT;
	BOOL enablePASS;
	BOOL enableFAIL;	
};

enum {					 HI_POT_EDIT = 0,   GROUND_BOND_EDIT,  POT_LOW_EDIT,  POT_HIGH_EDIT,	REMOTE_EDIT,    AC_SWEEP_EDIT,  SIGNAL_EDIT,      ACTUATOR_EDIT,  FILTER_ON_EDIT,  FILTER_OFF_EDIT,  RESULT_EDIT};
#endif