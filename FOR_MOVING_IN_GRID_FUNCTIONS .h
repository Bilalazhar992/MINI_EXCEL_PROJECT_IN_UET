#pragma once
#include<Windows.h>
#include <conio.h>
#include <iomanip>
#include <vector>
#include <string>
#include <fstream>
#define BLACK 0
#define BROWN 6
#define WHITE 15
#define GREEN 2
#define RED 4
#define LBLUE 9

void SetClr(int clr)
{
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), clr);
}

/*
* SetClr(int clr): This is a function that sets the text color of the console. It takes an integer clr as a parameter and uses the Windows API function SetConsoleTextAttribute to set the text color. The GetStdHandle(STD_OUTPUT_HANDLE) function is used to get the handle to the standard output (console).
*/
void SetClr_withbackgrd(int tcl, int bcl)
{
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), (tcl + (bcl * 16)));
}

/*
	SetClr_withbackgrd(int tcl, int bcl): This function sets both the text color and the background color of the console. It takes two integer parameters, tcl for the text color and bcl for the background color. It combines these colors and sets the text attribute using SetConsoleTextAttribute.
*/



/*
* gotoRowCol(int rpos, int cpos): This function moves the cursor to a specified row and column in the console window. It uses the SetConsoleCursorPosition function from the Windows API to set the cursor position to the specified row and column.
*/

void gotoRowCol(int rpos, int cpos)
{
	COORD scrn;
	HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
	scrn.X = cpos;
	scrn.Y = rpos;
	SetConsoleCursorPosition(hOuput, scrn);
}