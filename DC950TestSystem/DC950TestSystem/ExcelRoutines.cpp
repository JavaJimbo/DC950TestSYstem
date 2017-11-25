/*  ExcelRoutines.cpp - routines for creating a formatted spreadsheet and writing data to cells. 
 *	Requires files BasicExcel and ExcelFormat for low level spreadsheet functions.
 *
 */

/*
enum EXCEL_COLORS {
EGA_BLACK = 0,	// 000000H
EGA_WHITE = 1,	// FFFFFFH
EGA_RED = 2,	// FF0000H
EGA_GREEN = 3,	// 00FF00H
EGA_BLUE = 4,	// 0000FFH
EGA_YELLOW = 5,	// FFFF00H
EGA_MAGENTA = 6,	// FF00FFH
EGA_CYAN = 7		// 00FFFFH
};
*/

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

using namespace WinCompFiles;
using namespace YExcel;
using namespace ExcelFormat;
extern int currentRow;
char dateBuffer[80];

char ExcelFilename[MAXSTRING] = "C:\\DC950_Data\\DC950TestData.xls";

	const char* spreadsheetColTitles[TOTAL_COLUMNS + 1]	= {
		"TEST DATE",
		"SERIAL NUMBER",		
		"HI POT",
		"GROUND BOND ",
		"LAMP OFF",
		"LAMP ON",
		"5V REF",
		"REM VOLT 0",
		"REM VOLT 1",
		"REM VOLT 2",
		"REM VOLT 3",
		"REM VOLT 4",
		"REM VOLT 5",
		"INHIBIT",				
		"SWEEP 0",
		"SWEEP 1",
		"SWEEP 2",
		"SWEEP 3",
		"SWEEP 4",
		"SWEEP 5",		
		" TTL OUT ",
		"IRRADIANCE  ",
		"WAVELENGTH  ",
		"MIN AMPLITUDE",		
		" FWHM ",		
		"COLOR TEMP",		
		"IRRADIANCE OPEN",		
		"COLOR TEMP OPEN",		
		"RESULT",
		NULL
	}; 


// char ExcelFilename[MAXSTRING] = "";		
BasicExcel xls;


BOOL TestApp::createNewSpreadsheet()
{
	if (ExcelFilename == NULL) return FALSE;
		
	BasicExcelCell* cell = NULL;
	int col, row = 0;
	int textLength, columnWidth;

	ExcelFont font_magenta_bold;
	font_magenta_bold._weight = FW_BOLD;
	font_magenta_bold._color_index = EGA_MAGENTA;

	ExcelFont font_red_bold;
	font_red_bold._weight = FW_BOLD;
	font_red_bold._color_index = EGA_RED;

	ExcelFont font_blue_bold;
	font_blue_bold._weight = FW_BOLD;
	font_blue_bold._color_index = EGA_BLUE;

	ExcelFont font_black_bold;
	font_black_bold._weight = FW_BOLD;
	font_black_bold._color_index = EGA_BLACK;

	// create sheet 1 and get the associated BasicExcelWorksheet pointer
	xls.New(1);
	BasicExcelWorksheet* sheet = xls.GetWorksheet(0);

	XLSFormatManager fmt_mgr(xls);
		
	// TOP ROW:
	CellFormat fmt_black_bold(fmt_mgr, font_black_bold);
	fmt_black_bold.set_color1(COLOR1_PAT_SOLID);
	fmt_black_bold.set_color2(MAKE_COLOR2(EGA_YELLOW, 0));

	col = 0;
	row = 0;

	do {
		cell = sheet->Cell(row, col);
		textLength = (int)strlen(spreadsheetColTitles[col]);
		columnWidth = textLength * CHARWIDTH;
		sheet->SetColWidth(col, columnWidth);
		if (col == 0) cell->Set("DC950 TEST SYSTEM");
		cell->SetFormat(fmt_black_bold);
		col++;
	} while (col < 2);
		
		
	// TEST TITLE ROW: SERIAL NUMBER column: 	
	CellFormat fmt_blue_bold(fmt_mgr, font_blue_bold);
	fmt_blue_bold.set_color1(COLOR1_PAT_SOLID);
	fmt_blue_bold.set_color2(MAKE_COLOR2(EGA_GREEN, 0));

	
	col = 0;	
	do {
		row = 2;
		cell = sheet->Cell(row, col);
		textLength = (int)strlen(spreadsheetColTitles[col]);
		columnWidth = textLength * CHARWIDTH;
		sheet->SetColWidth(col, columnWidth);
		cell->Set(spreadsheetColTitles[col]);
		cell->SetFormat(fmt_blue_bold);

		row = 1;
		cell = sheet->Cell(row, col);
		cell->SetFormat(fmt_blue_bold);
		cell->Set("  ");

		col++;
	} while(col < HI_POT);

	// TEST TITLE ROW: do columns for all units
	col = HI_POT;
	row = 2;

	CellFormat fmt_red_bold(fmt_mgr, font_red_bold);
	fmt_red_bold.set_color1(COLOR1_PAT_SOLID);
	fmt_red_bold.set_color2(MAKE_COLOR2(EGA_CYAN, 0));

	do {
		cell = sheet->Cell(row, col);
		textLength = (int)strlen(spreadsheetColTitles[col]);
		columnWidth = textLength * CHARWIDTH;
		sheet->SetColWidth(col, columnWidth);
		cell->Set(spreadsheetColTitles[col]);   
		cell->SetFormat(fmt_red_bold);
		col++;
	} while (col < TTL_OUT);

	
	// SECOND ROW: "TESTS FOR ALL UNITS" TEXT
	fmt_blue_bold.set_color2(MAKE_COLOR2(EGA_CYAN, 0));
	row = 1;
	col = HI_POT;
	do {
		cell = sheet->Cell(row, col);
		if (col == HI_POT) cell->Set("'TESTS FOR ALL UNITS");		
		else cell->Set("");
		cell->SetFormat(fmt_blue_bold);
		col++;
	} while (col < TTL_OUT);

	// SECOND ROW: "TESTS FOR FILTER UNITS" TEXT
	fmt_blue_bold.set_color2(MAKE_COLOR2(EGA_YELLOW, 0));

	do {
		cell = sheet->Cell(row, col);
		if (col == TTL_OUT) cell->Set("'TESTS FOR FILTER UNITS");		
		else cell->Set("");
		cell->SetFormat(fmt_blue_bold);
		col++;
	} while (col < FINAL_RESULT);


	// TEST TITLE ROW: do columns for filter-actuator CLOSED
	row = 2;
	col = TTL_OUT;

	CellFormat fmt_magenta_bold(fmt_mgr, font_magenta_bold);
	fmt_magenta_bold.set_color1(COLOR1_PAT_SOLID);
	fmt_magenta_bold.set_color2(MAKE_COLOR2(EGA_YELLOW, 0));

	do {
		cell = sheet->Cell(row, col);
		textLength = (int)strlen(spreadsheetColTitles[col]);
		columnWidth = textLength * CHARWIDTH;
		sheet->SetColWidth(col, columnWidth);
		cell->Set(spreadsheetColTitles[col]);
		cell->SetFormat(fmt_magenta_bold);
		col++;
	} while (col < IRRADIANCE_OPEN);			

	// TEST TITLE ROW: do columns for filter-actuator OPEN		
	fmt_blue_bold.set_color2(MAKE_COLOR2(EGA_YELLOW, 0));

	do {
		cell = sheet->Cell(row, col);
		textLength = (int)strlen(spreadsheetColTitles[col]);
		columnWidth = textLength * CHARWIDTH;
		sheet->SetColWidth(col, columnWidth);
		cell->Set(spreadsheetColTitles[col]);
		cell->SetFormat(fmt_blue_bold);
		col++;
	} while (col < FINAL_RESULT);
	
	
	// TEST TITLE ROW: Do Final RESULT column: 
	fmt_blue_bold.set_color2(MAKE_COLOR2(EGA_GREEN, 0));
	col = FINAL_RESULT;
	row = 2;
	cell = sheet->Cell(row, col);
	textLength = (int)strlen(spreadsheetColTitles[col]);
	columnWidth = textLength * CHARWIDTH;
	sheet->SetColWidth(col, columnWidth);
	cell->Set(spreadsheetColTitles[col]);	
	cell->SetFormat(fmt_blue_bold);

	row = 1;
	cell = sheet->Cell(row, col);
	columnWidth = textLength * CHARWIDTH;
	sheet->SetColWidth(col, columnWidth);
	cell->Set("FINAL");
	cell->SetFormat(fmt_blue_bold);

	xls.SaveAs(ExcelFilename);
	return TRUE;
}


BOOL TestApp::CloseSpreadsheet() {	
	// xls.SaveAs(ExcelFilename);
	xls.Close();
	return TRUE;
}



BOOL TestApp::writeTestResultsToSpreadsheet(int col, int testResult, char *ptrTestData)
{		
	BasicExcelCell* cell = NULL;	
	BasicExcelWorksheet* sheet = xls.GetWorksheet(0);	

	XLSFormatManager fmt_mgr(xls);

	if (testResult == PASS) {
		ExcelFont font_black_bold;
		font_black_bold._weight = FW_BOLD;
		font_black_bold._color_index = EGA_BLACK;

		CellFormat fmt_result_bold(fmt_mgr, font_black_bold);
		fmt_result_bold.set_color1(COLOR1_PAT_SOLID);
		fmt_result_bold.set_color2(MAKE_COLOR2(EGA_WHITE, 0));

		cell = sheet->Cell(currentRow, col);
		cell->SetFormat(fmt_result_bold);

		if (ptrTestData != NULL) cell->Set(ptrTestData);
		else cell->Set("PASS");
	}		
	else {
		ExcelFont font_red_bold;
		font_red_bold._weight = FW_BOLD;
		font_red_bold._color_index = EGA_RED;

		CellFormat fmt_result_bold(fmt_mgr, font_red_bold);
		fmt_result_bold.set_color1(COLOR1_PAT_SOLID);
		fmt_result_bold.set_color2(MAKE_COLOR2(EGA_WHITE, 0));

		cell = sheet->Cell(currentRow, col);
		cell->SetFormat(fmt_result_bold);
		if (ptrTestData != NULL) cell->Set(ptrTestData);
		else cell->Set("FAIL");
	}	
	
	// saveAndCloseSpreadsheet();
	xls.SaveAs(ExcelFilename);
	return TRUE;

}

BOOL TestApp::loadSpreadsheet() {			
	xls.New(1);
	if (xls.Load(ExcelFilename)) {
		BasicExcelWorksheet *sheet1 = xls.GetWorksheet("Sheet1");		
	}
	else createNewSpreadsheet();
	
	xls.SaveAs(ExcelFilename);
	return TRUE;
}


BOOL TestApp::writeDateAndSerialNumberToSpreadsheet(char *ptrSerialNumber) {	
char dateBuffer[32];
struct tm  tstruct;		

	time_t now = time(0);
	tstruct = *localtime(&now);
	strftime(dateBuffer, sizeof(dateBuffer), "%m-%d-%Y", &tstruct);
	
	BasicExcelWorksheet* sheet = xls.GetWorksheet(0);
	currentRow = sheet->GetTotalRows();
	
	XLSFormatManager fmt_mgr(xls);	
	BasicExcelCell* cell = NULL;

	ExcelFont font_black_bold;
	font_black_bold._weight = FW_BOLD;
	font_black_bold._color_index = EGA_BLACK;

	CellFormat fmt_result_bold(fmt_mgr, font_black_bold);
	fmt_result_bold.set_color1(COLOR1_PAT_SOLID);
	fmt_result_bold.set_color2(MAKE_COLOR2(EGA_WHITE, 0));

	int col = TEST_DATE_COLUMN;
	cell = sheet->Cell(currentRow, col);
	cell->Set(dateBuffer);

	col = SERIAL_NUMBER_COLUMN;
	cell = sheet->Cell(currentRow, col);
	cell->Set(ptrSerialNumber);

	xls.SaveAs(ExcelFilename);
	return TRUE;
}

