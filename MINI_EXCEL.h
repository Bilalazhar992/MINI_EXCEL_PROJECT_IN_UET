#pragma once
#include <iostream>
#include <string.h>
#include "FOR_MOVING_IN_GRID_FUNCTIONS .h"
using namespace std;

static int cell_row_dim = 5;
static int cell_col_dim = 13;
static char symbol = 219;

class MiniExcel
{
private:
	class cell
	{

		string data; // data stored 
		cell* up;
		cell* down;
		cell* left;
		cell* right;
		int text_clr;
		char cell_alignment; // c for centre, r for right, l for left
		
		friend MiniExcel; // Mini_Excel is a friend class and it can use cell's private attributes
		//public:
		cell(string s = "", cell* u = nullptr, cell* d = nullptr, cell* l = nullptr, cell* r = nullptr, char cell_alignment = 'c')
		{
			data = s;
			up = u;
			down = d;
			left = l;
			right = r;
			text_clr = WHITE;
			cell_alignment = cell_alignment;
			
		}
	};
	int current_row, current_col; // row and column of the current cell
	cell head; // this is the 1st cell of the whole 2d matrix
	cell* head_pointer;
	cell* current_cell;
	cell* CellFromWhereRangeForOperationIsStarting;
	cell* CellFromWhereRangeForOperationIsEnding;
	// dimension of total cell grid on the screen (initially 5x5)
	int grid_row_dim;
	int grid_col_dim;
	bool is_a_cell_selected;
public:
	MiniExcel() 
	{
		cell head;
		this->head_pointer = &this->head;
		this->current_row = 0;
		this->current_col = 0;
		this->is_a_cell_selected = false;
		cout << "Do you want to load the previous sheet or want a new one (Y/N): ";
		char input;
		cin >> input;
		system("cls");
		this->current_cell = this->head_pointer;// 1st cell is the attribute of Mini_Excel class because we need a starting
		if (input == 'Y' || input == 'y')
		{
			int row_to_reach = 0, col_to_reach = 0;
			this->grid_row_dim = 1;
			this->grid_col_dim = 1;
			fstream rdr("E:\DSA\MINI EXCEL PROJECT\MINI EXCEL PROJECT",std::ifstream::in);
			rdr >> row_to_reach;
			rdr >> col_to_reach;

			int getline_index = 0;
			for (int i = 1; i < col_to_reach; i++)				
			{										
				this->InsertColumnAtRight(current_cell);
				current_cell = current_cell->right;
			}
			for (int j = 1; j < row_to_reach; j++)  
			{
				this->InsertRowBelow();
				current_cell = current_cell->down;
			}
			current_cell = this->head_pointer;
			string s;
			cell* goingDown = this->current_cell;
			cell* goingRight;
			getline(rdr, s);
			for (int i = 0; i < row_to_reach; i++)
			{
				getline(rdr, s);
				goingRight = goingDown;
				for (int j = 0; j < col_to_reach; j++)
				{
					string inp = "";
					while (true)
					{
						if (s[getline_index] == ' ' || s[getline_index] == '\n')
						{
							getline_index++;
							break;
						}
						inp.push_back(s[getline_index]);
						getline_index++;
					}
					if (inp != "_")
					{
						goingRight->data = inp;
					}
					goingRight = goingRight->right;// going right

				}
				getline_index = 0;
				goingDown = goingDown->down; // going down
			}
		}
		else
		{

			// cell
			this->grid_row_dim = 1;
			this->grid_col_dim = 1;

			for (int i = 1; i < 5; i++)				// we are going to initialize 4 cols
			{										// since 1 head cell is already present
				this->InsertColumnAtRight(current_cell);
				current_cell = current_cell->right; // moving right
			}
			for (int j = 1; j < 5; j++)  // A row is already formed in above loop as the 1st row
			{
				this->InsertRowBelow();
				current_cell = current_cell->down;// moving down
			}
			this->current_cell = this->head_pointer; // resetting the pointer to the start

		}
	}
	void InsertColumnAtRight(cell* current_cell)
	{
		cell* temp = current_cell;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		cell* StartCellSaver = temp;
		while (temp != nullptr)
		{
			InsertCellToRight(temp);
			temp = temp->down;
		}
		temp = StartCellSaver;
		while (temp != nullptr)
		{

			if (temp->up == nullptr && temp->down == nullptr)
			{
				//When We Have Only A Single Row in Our Grid
			}
			// if we are on the top of the column
			else if (temp->up == nullptr)
			{
				temp->right->down = temp->down->right;
			}// if we are at the end of column
			else if (temp->down == nullptr)
			{
				temp->right->up = temp->up->right;
			}// if we are in between
			else
			{
				temp->right->up = temp->up->right;
				temp->right->down = temp->down->right;
			}
			temp = temp->down;
		}
		this->grid_col_dim++;

	}
	void InsertRowBelow()
	{
		cell* saver = current_cell;// saving the cell at which we are right now
		// going at the start of the row in which our current cell is present
		while (saver->left != nullptr)
		{
			saver = saver->left;
		}
		cell* left_Cell = saver;
		while (saver != nullptr)
		{
			InsertCellDown(saver);
			saver = saver->right;
		}
		while (left_Cell != nullptr)
		{
			if (left_Cell->left == nullptr && left_Cell->right == nullptr)
			{
				//Special Case;
			}
			// we are at left most
			else if (left_Cell->left == nullptr)
			{
				left_Cell->down->right = left_Cell->right->down;
			}
			// we are at right most
			else if (left_Cell->right == nullptr)
			{
				left_Cell->down->left = left_Cell->left->down;
			} // we are in centre
			else
			{
				left_Cell->down->left = left_Cell->left->down;
				left_Cell->down->right = left_Cell->right->down;
			}
			left_Cell = left_Cell->right;
		}
		this->grid_row_dim++; // increasing the dimesion record if the grid
	}

	void InsertCellToRight(cell*& cur) // passed the current cell
	{						// cur = current_cell
		cell* helper = new cell;
		//cell* cur_temp = cur;

		// current cell is in the last column
		if (cur->right == nullptr)
		{		// a new cell with no up and down only left pointed to current cell
			cur->right = helper;
			helper->left = cur;
		}// current cell already has a cell on its right
		else if (cur->right != nullptr)
		{
			helper->right = cur->right;
			cur->right->left = helper;
			helper->left = cur;
			cur->right = helper;

		}
		// for now we are settign up and down of new cell to null 
		// and we will change it once we have constructed the whole column
		helper->up = nullptr;
		helper->down = nullptr;
	}
	void InsertCellDown(cell* cur)
	{
		cell* helper = new cell; // a new cell memory
		// if the current cell is on bottom 
		if (cur->down == nullptr)
		{
			cur->down = helper;
			helper->up = cur;
		} // agr centre ma kahin ha
		else
		{
			cell* temp = cur->down;
			helper->down=temp;
			helper->up = cur;
			temp->up = helper;
			cur->down = helper;
			
			
		}
		// initially setting left and right if the new cell equal to null
		helper->left = nullptr;
		helper->right = nullptr;
	}
	void StartExcel()
	{
		char keyBoardEvent;
		int range_start_col = 0, range_start_row = 0;
		int range_end_col = 0, range_end_row = 0;
		this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
		while (true)
		{
			this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
			keyBoardEvent = _getch(); // taking input from keyboard
			if (int(keyBoardEvent) == 72 || int(keyBoardEvent) == 80 || int(keyBoardEvent) == 75 || int(keyBoardEvent) == 77)
			{
				
				current_cell_mover(keyBoardEvent);

			}// if the user has selected a cell to type in it
			else if (keyBoardEvent == 'R' || keyBoardEvent == 'r')//Selecting Cell for Typing in it
			{
				this->is_a_cell_selected = true;
				while (is_a_cell_selected)
				{
					keyBoardEvent = _getch();  // taking input for what is going to be written in the cell
					if (keyBoardEvent == 'R' || keyBoardEvent == 'r')
					{
						this->is_a_cell_selected = false;
						break; // this means come out of the cell selected mode
					}
					else  // somthing is about to enter the cell data
					{     // if the input is valid
						if (is_input_data_valid(this->current_cell, keyBoardEvent))
						{
							enter_and_display_data_in_cell(this->current_cell, this->current_row, this->current_col, keyBoardEvent);
						}
					}

				}

			}  // agr column ya row add karna ka kaha ha add 
			else if (keyBoardEvent == 'I' || keyBoardEvent == 'i')  // i stands for insert
			{
				keyBoardEvent = _getch();
				if (keyBoardEvent == 'R' || keyBoardEvent == 'r') // r stands for row
				{
					keyBoardEvent = _getch();
					if (this->grid_row_dim <= 19)  // max rows that can be inserted are 20
					{
						if (keyBoardEvent == 'A' || keyBoardEvent == 'a') // a stands for above
						{

							this->InsertRowAbove();
							system("cls");
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyBoardEvent == 'B' || keyBoardEvent == 'b') // b stands for below
						{
							this->InsertRowBelow();
							system("cls");
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}

				}
				else if (keyBoardEvent == 'C' || keyBoardEvent == 'c')  // c stands for column
				{
					keyBoardEvent = _getch();
					if (this->grid_col_dim <= 9)
					{
						if (keyBoardEvent == 'R' || keyBoardEvent == 'r') // r stands for right
						{
							this->InsertColumnAtRight(current_cell);
							system("cls");
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyBoardEvent == 'L' || keyBoardEvent == 'l') // l stands for left
						{
							this->InsertColumnAtLeft(this->current_cell);
							system("cls");
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
				}
			} // if we want to delete  a column or a row
			else if (keyBoardEvent == 'D' || keyBoardEvent == 'd')
			{
				keyBoardEvent = _getch();
				if (keyBoardEvent == 'R' || keyBoardEvent == 'r') // r stands for row
				{

					if (this->grid_row_dim > 3)
					{
						this->delete_current_row();
						system("cls");
						this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}

				}
				else if (keyBoardEvent == 'C' || keyBoardEvent == 'c')  // c stands for column
				{

					if (this->grid_col_dim > 3)
					{
						this->delete_current_column();
						system("cls");
						this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
				}
			}
			else if (keyBoardEvent == 'c' || keyBoardEvent == 'C') // cell change mode  (like string alligment, color)
			{
				int pr = 70, pc = 70;
				keyBoardEvent = _getch();
				if (keyBoardEvent == 'k' || keyBoardEvent == 'K')  // k is for color change
				{

					gotoRowCol(pr, pc);
					char klr;
					SetClr(15);
					cout << "Enter the colour of the cell text you want, options:(White:w , green:g , blue:b) : ";
					klr = _getch();;
					while (true)
					{
						if (klr == 'w' || klr == 'W' || klr == 'g' || klr == 'G' || klr == 'b' || klr == 'B')
						{
							break;
						}
						gotoRowCol(pr, pc);
						cout << "                                                                               ";
						gotoRowCol(pr, pc);
						cout << "You have entered invalid colour please enter again: ";
						klr = _getch();
					}
					gotoRowCol(pr, pc);
					cout << "                                                                                   ";
					if (klr == 'w' || klr == 'W')
					{
						this->current_cell->text_clr = WHITE;
					}
					else if (klr == 'g' || klr == 'G')
					{
						this->current_cell->text_clr = GREEN;
					}
					else if (klr == 'b' || klr == 'B')
					{
						this->current_cell->text_clr = LBLUE;
					}
				}
				else if (keyBoardEvent == 's' || keyBoardEvent == 'S') // s for style change of the text
				{
					gotoRowCol(pr, pc);
					char klr;
					SetClr(15);
					cout << "Enter the Alignment type of text you want (c: centre, r: right, l:left): ";
					klr = _getch();;
					while (true)
					{
						if (klr == 'R' || klr == 'r' || klr == 'c' || klr == 'C' || klr == 'L' || klr == 'l')
						{
							break;
						}
						gotoRowCol(pr, pc);
						cout << "                                                                                          ";
						gotoRowCol(100, 115);
						cout << "You have entered invalid alignnmet type please enter again: ";
						klr = _getch();;
					}
					gotoRowCol(pr, pc);
					cout << "                                                                                                 ";
					if (klr == 'r' || klr == 'R')
					{
						this->current_cell->cell_alignment = 'r';
					}
					else if (klr == 'c' || klr == 'C')
					{
						this->current_cell->cell_alignment = 'c';
					}
					else if (klr == 'L' || klr == 'l')
					{
						this->current_cell->cell_alignment = 'l';
					}
				}
			}

			//++++++++++++++  E stands for erase command (this will be used to erase the data of a column or row) +++++++++++++

			else if (keyBoardEvent == 'e' || keyBoardEvent == 'E')
			{
				keyBoardEvent = _getch();
				if (keyBoardEvent == 'c' || keyBoardEvent == 'C')
				{
					clear_whole_column();
					this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}
				else if (keyBoardEvent == 'r' || keyBoardEvent == 'R')
				{
					clear_whole_row();
					this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}
			}


			//++++++++++++++++++++++++++++++++++   insert 1 (cell)    ++++++++++++++++++++++++++++++++++++++++++++++++++=

			else if (keyBoardEvent == '1')
			{
				keyBoardEvent = _getch();
				// inserting in column side
				if (keyBoardEvent == 'i' || keyBoardEvent == 'I')
				{
					keyBoardEvent = _getch();
					if (this->grid_col_dim <= 9)
					{
						if (keyBoardEvent == 'r' || keyBoardEvent == 'R')  // 1 cell at right with right shift
						{
							this->insert_single_cell_by_right_shift();
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyBoardEvent == 'l' || keyBoardEvent == 'L')
						{
							this->insert_single_cell_by_left_shift();
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
					if (this->grid_row_dim <= 19)
					{
						if (keyBoardEvent == 'u' || keyBoardEvent == 'U')   // 1 cell down with up shift
						{
							this->insert_single_cell_by_up_shift();
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
						else if (keyBoardEvent == 'D' || keyBoardEvent == 'd') // 1 cell up with down shift
						{
							this->insert_single_cell_by_down_shift();
							this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
						}
					}
				}


				// deleting a single cell

			}
			else if (keyBoardEvent == 'N' || keyBoardEvent == 'n')
			{
				//++++++++++=    insert and check the single cell deletion functions
				keyBoardEvent = _getch();



				if (keyBoardEvent == 'l' || keyBoardEvent == 'L')
				{
					this->delete_cell_by_leftShift();

					this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}

				else if (keyBoardEvent == 'u' || keyBoardEvent == 'U')  // 1 cell up with up shift
				{
					this->delete_cell_by_UpShift();
					this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}

			}

			else if (keyBoardEvent == 'm' || keyBoardEvent == 'M')
			{	// agr koi range select hoi ha too proceed karo warna terminate karka bahir aa jao

				if (GiveRange(range_start_row, range_start_col, range_end_col, range_end_row))
				{
					
					vector<float> value = GetInputForApplyingFunction(range_start_row, range_start_col, range_end_col, range_end_row);
					float result = 0;

					string input;
					while (true)
					{
						gotoRowCol(200, 0);
						cout << "Enter Which Opertion You want to Perform:   ";
						cin >> input;
						if (input == "Sum")
						{
							result = this->Sumation(value);
							break;
						}
						else if (input == "Count")
						{
							result = float(this->Count(value));
							break;
						}
						else if (input == "Average")
						{
							result = this->Average(value);
							break;
						}
						else if (input == "Minimum")
						{
							result = this->Minimum(value);
							break;
						}
						else if (input == "Maximum")
						{
							result = this->Maximum(value);
							break;
						}
						else
						{
							cout << "This Function is Not Available In Our Program!  ";
							gotoRowCol(200, 0);
							cout << "                                                                    " << endl;
							cout << "                                                                    ";

						}
					}
					cout << result;
					cout << "\nAny Key To Proceed:  ";
					_getch();
					system("cls");
					this->ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					// discard 	
					value.clear();

				}

			}

			// ++++++++++++     copy/ cut paste
			else if (keyBoardEvent == 'p' || keyBoardEvent == 'P' || keyBoardEvent == 88 || keyBoardEvent == 120) //  ctrl + (c=3  and   x=4)
			{    // if range is selected then 
				if (GiveRange(range_start_row, range_start_col, range_end_col, range_end_row))
				{
					if (keyBoardEvent == 'p' || keyBoardEvent == 'P') // copy 
					{
						CopyOrCut(range_start_row, range_start_col, range_end_row, range_end_col, false);
						ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
					else // cut
					{
						CopyOrCut(range_start_row, range_start_col, range_end_row, range_end_col, true);
						ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
					}
				}
			}
			
			else if (keyBoardEvent == 's' || keyBoardEvent == 'S')//Cell Swap
			{
				if (GiveRange(range_start_row, range_start_col, range_end_col, range_end_row))
				{
					SwapTwoCells(CellFromWhereRangeForOperationIsStarting);
					ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
				}
				
			}
			else if (keyBoardEvent == 'b' || keyBoardEvent == 'B')
			{
				SaveExcelSheet();
			}

		
		
		}
	}
	void SwapTwoCells(cell* temp)
	{
		string data = this->current_cell->data;
		this->current_cell->data=temp->data;
		temp->data = data;
	}

	float Sumation(vector<float>values)
	{
		float sum = 0;
		for (auto i : values)
		{
			sum += i;
		}
		return sum;
	}

	float Average(const vector<float>& values)
	{
		float avg = Sumation(values);
		avg /= values.size();
		return avg;
	}

	int Count(const vector<float>& values)
	{
		return values.size();
	}

	float Minimum(const vector<float>& values)
	{
		auto i = min_element(values.begin(), values.end());
		return float(*i);
	}

	float Maximum(const vector<float>& values)
	{
		auto i = max_element(values.begin(), values.end());
		return float(*i);
	}

	void delete_cell_by_UpShift()
	{
		cell* temp = current_cell;
		if (temp->down != nullptr)
		{
			while (temp->down != nullptr)
			{
				temp->data = temp->down->data;
				temp = temp->down;
			}

		}
		temp->data.clear();
		if (current_cell->up != nullptr)
		{
			current_cell_mover(72);
		}
	}

	void delete_cell_by_leftShift()
	{
		cell* temp = current_cell;
		if (temp->right == nullptr) // agr right pa cell hi nhi too delete kia karna 
		{
			temp->data.clear();

		}
		else
		{
			while (temp->right != nullptr)
			{
				temp->data = temp->right->data;
				temp = temp->right;
			}
			temp->data.clear();
		}
		if (current_cell->left != nullptr)
		{
			current_cell_mover(75);
		}


	}

	void CopyOrCut(int rangeStartRow, int rangStartCol, int rangeEndrow, int rangeEndColumn, bool Cut_Allow)
	{
		char keyBoardResponse;
		vector<string> CollectedData;

		//*************************************************
		 // for same row (horizontal case only)
		//*************************************************

		if (rangeStartRow == rangeEndrow)
		{
			if (rangStartCol > rangeEndColumn)
			{
				swap(this-> CellFromWhereRangeForOperationIsStarting, this->CellFromWhereRangeForOperationIsEnding);
				swap(rangStartCol, rangeEndColumn);
			}
			cell* temp = CellFromWhereRangeForOperationIsStarting;
			// we are producing this copy only if after performing copy or cut operation on some range we also want to perform some arithematic operation on this range also
			while (CellFromWhereRangeForOperationIsStarting != CellFromWhereRangeForOperationIsEnding)
			{
				CollectedData.push_back(CellFromWhereRangeForOperationIsStarting->data);
				CellFromWhereRangeForOperationIsStarting = CellFromWhereRangeForOperationIsStarting->right;
			}
			CollectedData.push_back(CellFromWhereRangeForOperationIsStarting->data);
			CellFromWhereRangeForOperationIsStarting = temp;
			if (Cut_Allow)
			{
				while (temp != CellFromWhereRangeForOperationIsEnding)
				{
					temp -> data.clear();
					temp = temp->right;

				}
				temp->data.clear();
				ExcelSheetPrinter(this->head_pointer, grid_row_dim, grid_col_dim);
				this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
			}
			while (true)
			{
				keyBoardResponse = _getch();
				if (int(keyBoardResponse) == 72 || int(keyBoardResponse) == 80 || int(keyBoardResponse) == 75 || int(keyBoardResponse) == 77)
				{
					
					current_cell_mover(keyBoardResponse);
					this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
				}
				else if (int(keyBoardResponse) == 86 || int(keyBoardResponse) == 118)
				{
					cell* temp = current_cell;
					
					int CurrentColCopy = current_col;
					for (int n = 0;n < CollectedData.size();n++)
					{
						temp->data = CollectedData[n];
						
						if (CurrentColCopy <= 9 && temp->right == nullptr)
						{
							//ExcelSheetPrinter(this->head_pointer, this->grid_row_dim, this->grid_col_dim);
							this->InsertColumnAtRight(temp);
							
							
							
						}
						else if(CurrentColCopy > 9)
						{
							break;
						}
						
						
						temp = temp->right;
						CurrentColCopy++;
						current_cell_mover(77);
						
						
					}
					delete_current_column();
					this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
					
					
					break;
					
				}
			}
			CollectedData.clear();
			return;
		}

		//*************************************************
		 // for same column (vertical case only)
		//*************************************************

		else if (rangStartCol == rangeEndColumn)
		{
			if (rangeStartRow > rangeEndrow)
			{
				swap(this->CellFromWhereRangeForOperationIsStarting, this->CellFromWhereRangeForOperationIsEnding);
				swap(rangeStartRow, rangeEndrow);
			}
			cell* temp = CellFromWhereRangeForOperationIsStarting;
			// we are producing this copy only if after performing copy or cut operation on some range we also want to perform some arithematic operation on this range also
			while (CellFromWhereRangeForOperationIsStarting != CellFromWhereRangeForOperationIsEnding)
			{
				CollectedData.push_back(CellFromWhereRangeForOperationIsStarting->data);
				CellFromWhereRangeForOperationIsStarting = CellFromWhereRangeForOperationIsStarting->down;
			}
			CollectedData.push_back(CellFromWhereRangeForOperationIsStarting->data);
			CellFromWhereRangeForOperationIsStarting = temp;
			if (Cut_Allow)
			{
				while (temp != CellFromWhereRangeForOperationIsEnding)
				{
					temp->data.clear();
					temp = temp->down;

				}
				temp->data.clear();
				ExcelSheetPrinter(this->head_pointer, grid_row_dim, grid_col_dim);
				this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
			}
			while (true)
			{
				keyBoardResponse = _getch();
				if (int(keyBoardResponse) == 72 || int(keyBoardResponse) == 80 || int(keyBoardResponse) == 75 || int(keyBoardResponse) == 77)
				{

					current_cell_mover(keyBoardResponse);
					this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
				}
				else if (int(keyBoardResponse) == 86 || int(keyBoardResponse) == 118)
				{
					
					int CurrentRowCopy = current_row;
					for (int n = 0;n < CollectedData.size();n++)
					{
						current_cell->data = CollectedData[n];
						if (CurrentRowCopy <= 19 && current_cell->down == nullptr)
						{
							this->InsertRowBelow();
						}
						else if (CurrentRowCopy > 19)
						{
							break;
						}
						
						CurrentRowCopy++;
						current_cell_mover(80);
						if (n != CollectedData.size() - 1)
						{
							this->current_cell_printer(this->current_cell, this->current_row, this->current_col);

						}

					}
					
					
					
					break;
				}
				
			}
			delete_current_row();
			CollectedData.clear();
			return;
		}

	}

	// saving data on a new file
	void SaveExcelSheet()
	{
		fstream write("E:\DSA\MINI EXCEL PROJECT\MINI EXCEL PROJECT", std::ofstream::out|std::ofstream::trunc);
		write << this->grid_row_dim << " " << this->grid_col_dim << endl;
		cell* ForMovingDown=this->head_pointer;
		cell* ForMovingStartOfRow;
		for (int i = 0; i < this->grid_row_dim; i++)
		{
			ForMovingStartOfRow= ForMovingDown;
			for (int j = 0; j < this->grid_col_dim; j++)
			{
				if (ForMovingStartOfRow->data == "")
				{
					write << "_" << " ";  // if the cell was empty
				}
				else
				{
					write << ForMovingStartOfRow->data << " ";
				}
				ForMovingStartOfRow= ForMovingStartOfRow->right;
			}
			write << endl;
			ForMovingDown = ForMovingDown->down;
		}
	}

	void insert_single_cell_by_down_shift()
	{
		cell* saver = this->current_cell;
		while (this->current_cell->down != nullptr)
		{
			this->current_cell = this->current_cell->down;  // going down
		}
		// now using already existing column insert function
		this->InsertRowBelow(); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		cell* down_most = this->current_cell;
		down_most = down_most->down;
		this->current_cell = saver; // re position the current pointer to its original position

		// now shifting the data of cells into the cell 
		while (down_most != this->current_cell)
		{
			down_most->data = down_most->up->data; // shifting to right
			down_most = down_most->up; // going up
		}
		down_most->data.clear();
		
	}


	void insert_single_cell_by_right_shift()
	{
		// saving current cell
		cell* saver = this->current_cell;
		while (saver->right != nullptr)
		{
			saver = saver->right;  // going right
		}
		// now using already existing column insert function
		this->InsertColumnAtRight(saver); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		cell* right_most = saver;


		// now shifting the data of cells into the cell 
		while (right_most != this->current_cell)
		{
			right_most->right->data = right_most->data; // shifting to right
			right_most = right_most->left; // going left
		}
		right_most->right->data.clear();
		

	}

	void insert_single_cell_by_left_shift()
	{
		// saving current cell
		cell* saver = this->current_cell;
		while (saver->left != nullptr)
		{
			saver = saver->left;  // going left
		}
		// now using already existing column insert function
		this->InsertColumnAtLeft(saver); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		cell* left_most = saver;
		left_most = left_most->left;


		// now shifting the data of cells into the cell 
		while (left_most != this->current_cell)
		{
			left_most->data = left_most->right->data; // shifting to right
			left_most = left_most->right; // going right
		}
		left_most->data.clear();
		//current_cell_mover(77);
		//this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
	}


	void insert_single_cell_by_up_shift()
	{
		cell* saver = this->current_cell;
		while (this->current_cell->up != nullptr)
		{
			this->current_cell = this->current_cell->up;  // going up
		}
		// now using already existing column insert function
		this->InsertRowAbove(); // this uses current iterator 

		// making an iterator pointing at the current cell which is actually the cell before the last cell of the row 
		cell* up_most = this->current_cell;
		up_most = up_most->up;
		this->current_cell = saver; // re position the current pointer to its original position


		// now shifting the data of cells into the cell 
		while (up_most != this->current_cell)
		{
			up_most->data = up_most->down->data; // shifting to right
			up_most = up_most->down; // going down
		}
		up_most->data.clear();
		

	}



	void clear_whole_row()
	{
		cell* copy = this->current_cell;
		// going at the start of row;
		while (copy->left != nullptr)
		{
			copy = copy->left;  // going up
		}
		// now going right and also clearing the whole row
		while (copy != nullptr)
		{
			copy->data.clear();
			copy = copy->left; // going down
		}
	}
	void clear_whole_column()
	{
		cell* copy = this->current_cell;
		// going at the start of column;
		while (copy->up != nullptr)
		{
			copy = copy->up;  // going up
		}
		// now going down and also clearing the whole column
		while (copy != nullptr)
		{
			copy->data.clear();
			copy = copy->down; // going down
		}
	}

	void current_cell_printer(cell* curr_cell, int cur_cell_row, int cur_cell_col)
	{
		this->print_cell(curr_cell, RED, cur_cell_row, cur_cell_col);
	}

	//++++  function for moving the current_cell in the direction  of the arrow key pressed
	void current_cell_mover(char arrow_key)
	{
		if (arrow_key == 72) // up arrow key pressed
		{
			if (this->current_cell->up != nullptr)
			{
				print_cell(current_cell, WHITE, current_row, current_col);
				this->current_cell = this->current_cell->up;
				this->current_row--;
			}
		}
		else if (arrow_key == 80) // down arrow key pressed
		{
			if (this->current_cell->down != nullptr)
			{
				print_cell(current_cell, WHITE, current_row, current_col);
				this->current_cell = this->current_cell->down;
				this->current_row++;
			}
		}
		else if (arrow_key == 77) // right arrow key pressed
		{
			if (this->current_cell->right != nullptr)
			{
				print_cell(current_cell, WHITE, current_row, current_col);
				this->current_cell = this->current_cell->right;
				this->current_col++;
			}
		}
		else if (arrow_key == 75) // left arrow key pressed
		{
			if (this->current_cell->left != nullptr)
			{
				print_cell(current_cell, WHITE, current_row, current_col);
				this->current_cell = this->current_cell->left;
				this->current_col--;
			}
		}
	}

	void enter_and_display_data_in_cell(cell* current_cell, int current_row, int current_col, char x)
	{
		if (int(x) != 8) // if x is not equal to backspace
		{
			current_cell->data.push_back(x);
		}
		else
		{
			if (!current_cell->data.empty()) // agr backspace press hoa ha or data ha cell ma too ak char delete kar do
			{
				current_cell->data.pop_back();
			}
		}
		print_data_in_cell(current_cell, current_row, current_col);
	}


	bool is_input_data_valid(cell* c, char x)
	{
		if (int(x) == -32)
		{
			x = _getch(); // if we have pressed arrow keys then we have to also waste its both characters
			return false;
		}

		if (int(x) == 8) // backspace char 
		{
			return true;
		}

		if (!((int(x) >= 48 && int(x) <= 57) || ((x >= 'A' || x >= 'a') && (x <= 'z' || x <= 'Z')) || x == '.')) // agr words, alpha bets or decimal ka alava ha too enter nhi karna
		{
			return false; // agr integers or alpha bets ka alava koi chz ha too wo wsa hi enter nhi ho gi
		}
		// agr abhi cell ha hi empty too direct put kar doo
		if (c->data.empty())
		{
			return true;
		}
		else if (c->data.size() == 4) // agr wo limit pa ha too put nhi karna
		{
			return false;
		}
		else if (x == '.') // agr input decimal aya ha or pahla bhi ak decimal ha too usa new nahin dalna chia
		{
			for (int k{}; k < c->data.size(); k++)
			{
				if (c->data[k] == '.' || ((c->data[0] >= 'A' || c->data[0] >= 'a') && (c->data[0] <= 'z' || c->data[0] <= 'Z')))
				{
					return false;
				}
			}
		}
		else if ((int(c->data[0]) >= 48 && int(c->data[0]) <= 57) || c->data[0] == '.') // agr string ka pahla char number ha ya phir decimal ha
		{
			if ((x >= 'A' || x >= 'a') && (x <= 'z' || x <= 'Z')) // or jo char put karna lga han wo alphabet ha
			{
				// then it cannot be done
				return false;
			}

		}
		else if ((c->data[0] >= 'A' || c->data[0] >= 'a') && (c->data[0] <= 'z' || c->data[0] <= 'Z')) // agr string ka pahla char alphabet ha
		{
			if ((int(x) >= 48 && int(x) <= 57) || x == '.') // or tum number put karna lag ho ya decimal wo bhi allow nhi ha
			{
				return false;
			}
		}
		return true;
	}
	//____  function for entring data in a cell





	void ExcelSheetPrinter(cell* head_poniter, int grid_total_row, int grid_total_col) // making no copy
	{

		
		//++++ Now Printing the grid
		//  This is going to be a zig zag printing 

		
		cell* printing_head = this->head_pointer; // iterator at the start of the grid
		for (int i = 0; i < grid_total_row; i++)
		{
			if (i % 2 == 0) // if i is even
			{
				for (int j = 0; j < grid_total_col; j++)
				{
					print_cell(printing_head, WHITE, i, j);
					if (j < grid_total_col - 1) // because if we have 2 cells then we will prin for 0 and 1 if we ++ after that the printing head will become null
					{
						printing_head = printing_head->right; // moving from most left to most right
					}

				}
				printing_head = printing_head->down; // moving one row down
			}
			else
			{
				for (int j = grid_total_col - 1; j >= 0; j--)
				{
					print_cell(printing_head, WHITE, i, j);
					if (j > 0)
					{
						printing_head = printing_head->left; // moving from most right to most left
					}
				}
				printing_head = printing_head->down; // moving one row down
			}
			// Note: at the end the printing head iterator will become null
		}


	}
	// single cell printer
	void print_cell(cell* which_cell, int clr, int cur_row, int cur_col)
	{
		SetClr(clr);
		int strt_r = cur_row * cell_row_dim, strt_c = cur_col * cell_col_dim;
		gotoRowCol(cur_row * cell_row_dim, cur_col * cell_col_dim);
		for (int i = 0; i < cell_row_dim; i++)
		{
			for (int j = 0; j < cell_col_dim; j++)
			{
				gotoRowCol(strt_r + i, strt_c + j);
				if (i == 0 || j == 0 || i == cell_row_dim - 1 || j == cell_col_dim - 1)
				{
					cout << symbol;
				}

			}
		}
		SetClr(15);
		print_data_in_cell(which_cell, cur_row, cur_col);
	}
	void print_data_in_cell(cell* receive, int which_row, int which_col)
	{
		cell* c = receive;
		clear_data_in_cell(which_row, which_col);
		if (c->cell_alignment == 'l' || c->cell_alignment == 'L')  // left allignmnet
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 2);
		}
		else if (c->cell_alignment == 'r' || c->cell_alignment == 'R') // right allignment
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 7);
		}
		else // centre allignment
		{
			gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 5);
		}
		SetClr(c->text_clr);
		for (int i = 0; i < c->data.size(); i++)
		{
			cout << c->data[i]; // max it will contain 3 characters
		}
		SetClr(15); // back to white
	}

	void clear_data_in_cell(int which_row, int which_col)
	{
		gotoRowCol((which_row * cell_row_dim) + 2, (which_col * cell_col_dim) + 2);
		cout << "          ";
	}


	void InsertRowAbove()
	{
		bool fl = false;
		cell* temp = this->current_cell; // saving the cell at which we are right now
		// going at the start of the row in which our current cell is present
		while (temp->left != nullptr)
		{
			temp = temp->left; // going left
		}
		cell* left_most = temp;// left cell of the row ko save kar rha hon
		if (left_most == this->head_pointer)
		{
			fl = true;
		}
		///++++++    inserting the cells
		while (temp != nullptr) // hum na last cell ka uper bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellUp(temp);
			temp = temp->right; // moving right in row
		}
		if (fl)  // if cell is placed above the head pointr then reassign the head pointer to it
		{
			head_pointer = left_most->up;
		}

		// now changing the configuration of the cells we just inserted
		while (left_most != nullptr)
		{
			if (left_most->left == nullptr && left_most->right == nullptr)
			{
				// do nothing
			}
			// we are at left most
			else if (left_most->left == nullptr /*&& left_most->right != nullptr*/)
			{
				left_most->up->right = left_most->right->up;
			}
			// we are at right most
			else if (/*left_most->left != nullptr &&*/ left_most->right == nullptr)
			{
				left_most->up->left = left_most->left->up;
			} // we are in centre
			else
			{
				left_most->up->left = left_most->left->up;
				left_most->up->right = left_most->right->up;
			}
			left_most = left_most->right; // going from left to right
		}
		this->current_row++; 
		this->grid_row_dim++;
	}

	void InsertCellUp(cell* cur)  // cur = current cell
	{
		cell* helper = new cell; // a new cell memory
		// if the current cell is on top 
		if (cur->up == nullptr)
		{
			cur->up = helper;
			helper->down = cur;
		} // agr centre ma kahin ha
		else
		{
			cur->up->down = helper;
			helper->up = cur->up;
			helper->down = cur;
			cur->up = helper;
		}
		// initially setting left and right if the new cell equal to null
		helper->left = nullptr;
		helper->right = nullptr;

	}
	void InsertColumnAtLeft(cell* upperSyAya)
	{
		bool fl = false;
		cell* temp = upperSyAya; // saving the cell at which we are right now
		// going at the start of the column in which our current cell is present
		while (temp->up != nullptr)
		{
			temp = temp->up; // going up
		}
		cell* TopAdressSaver = temp;// top cell of the column ko save kar rha hon
		if (temp == this->head_pointer)  // agr us column ka left pa ak new column insert karna ha jis ma head majood ha too head ko bhi dobara assign karna para ga 
		{
			fl = true;
		}
		while (temp != nullptr) // hum na last cell ka aga bhi insert karna ha or us sa aga ya null ho jai ga
		{
			InsertCellToLeft(temp);
			temp = temp->down; // moving dowm in column
		}
		if (fl)
		{    // reassigning of head if it is pushed to right
			this->head_pointer = TopAdressSaver->left;
		}
		// now changing the up and down pointers of the new column we just inserted
		while (TopAdressSaver != nullptr)
		{
			if (TopAdressSaver->up == nullptr && TopAdressSaver->down == nullptr)
			{
				//continue;
				// do nothing
			}
			// if we are on the top of column
			else if (TopAdressSaver->up == nullptr)
			{
				TopAdressSaver->left->down = TopAdressSaver->down->left;

			}  // if we are at the end of column
			else if (TopAdressSaver->down == nullptr)
			{
				TopAdressSaver->left->up = TopAdressSaver->up->left;

			} // if we are in between
			else
			{
				TopAdressSaver->left->up = TopAdressSaver->up->left;
				TopAdressSaver->left->down = TopAdressSaver->down->left;
			}
			TopAdressSaver = TopAdressSaver->down;
		}
		this->current_col++; // since current cell ka picha ak column add hoa ha is lia us ka column number ak aga chala gia ha
		this->grid_col_dim++;
	}

	void InsertCellToLeft(cell* cur)
	{
		cell* helper = new cell;
		//cell* cur_temp = cur;

		// current cell is in the 1st column
		if (cur->left == nullptr)
		{		// a new cell with no up and down only right pointed to current cell
			cur->left = helper;
			helper->right = cur;
		}// current cell already has a cell on its left
		else
		{
			cell* temp = cur->left;
			helper->right = cur;
			helper->left = temp;
			temp->right = helper;
			cur->left = helper;

		}
		// for now we are settign up and down of new cell to null 
		// and we will change it once we have constructed the whole column	
		helper->up = nullptr;
		helper->down = nullptr;
	}

	void delete_current_row()
	{
		cell* temp = this->current_cell;
		while (temp->left != nullptr)
		{
			temp = temp->left; // going left
		}
		if (temp == this->head_pointer)  // agr jo row delete karna laha han wo head pa ha too head ko bhi shift kar doo
		{
			this->head_pointer = this->head_pointer->down; // shifting to down cell
		}
		// shifting the current cell
		if (this->current_cell->down == nullptr)
		{
			this->current_cell = this->current_cell->up;
			this->current_row--;
		}
		else
		{
			this->current_cell = this->current_cell->down;

		}
		//  Now deleting the whole row
		while (temp != nullptr)
		{
			if (temp->up != nullptr)
			{
				temp->up->down = temp->down;
			}
			if (temp->down != nullptr)
			{
				temp->down->up = temp->up;
			}
			temp->up = nullptr; temp->down = nullptr;
			// if we are in the middle and we have to 1st shift the helper to its right
			if (temp->right != nullptr)
			{
				cell* del = temp;
				temp = temp->right;
				temp->left = nullptr;
				del->left = nullptr; del->right = nullptr;
				//delete[]del;
			} // if we are already at the end then we are not going to shift and only delete helper
			else
			{
				temp->left = nullptr;
				//delete[]temp;
				break;
			}

		}
		this->grid_row_dim--;
	}

	void delete_current_column()
	{
		cell* temp = this->current_cell;
		while (temp->up != nullptr)
		{
			temp = temp->up; // going up
		}
		if (temp == this->head_pointer)  // agr jo row delete karna laha han wo head pa ha too head ko bhi shift kar doo
		{
			this->head_pointer = this->head_pointer->right; // shifting to right cell
		}
		// shifting the current cell
		if (this->current_cell->right == nullptr)
		{
			this->current_cell = this->current_cell->left;
			this->current_col--;
		}
		else
		{
			this->current_cell = this->current_cell->right;
		}
		//  Now deleting the whole column
		while (temp != nullptr)
		{
			if (temp->left != nullptr)
			{
				temp->left->right = temp->right;
			}
			if (temp->right != nullptr)
			{
				temp->right->left = temp->left;
			}
			temp->left = nullptr; temp->right = nullptr;
			// if we are in the middle and we have to 1st shift the helper to its down
			if (temp->down != nullptr)
			{
				cell* del = temp;
				temp = temp->down; // going down
				temp->up = nullptr;
				del->down = nullptr; del->up = nullptr;
				//delete[]del;
			}
			// if we are already at the end then we are not going to shift and only delete helper
			else
			{
				temp->up = nullptr;
				//delete[]temp;
				break;
			}

		}
		this->grid_col_dim--;
	}

	bool GiveRange(int& range_start_row, int& range_start_col, int& range_end_col, int& range_end_row)
	{
		char keyBoardEvent;
		range_start_row = this->current_row, range_start_col = this->current_col;
		this->CellFromWhereRangeForOperationIsStarting = current_cell;
		while (true)  // end the range
		{
			keyBoardEvent = _getch();
			// 1st we are going to select a range
			if (int(keyBoardEvent) == -32)
			{
				keyBoardEvent = _getch();
				current_cell_mover(keyBoardEvent);
				this->current_cell_printer(this->current_cell, this->current_row, this->current_col);
			}
			else if (keyBoardEvent == 'E' || keyBoardEvent == 'e')  // end range function and perform the function on the selected range
			{
				range_end_row = this->current_row; range_end_col = this->current_col;
				this->CellFromWhereRangeForOperationIsEnding = this->current_cell;
				return true;// continue the range function
			}

		}
	}


	void value_inserter_for_range_function(cell* x, vector<float>& values)
	{
		string val_to_check = x->data;
		if (!val_to_check.empty())
		{
			if (!is_alphabet((val_to_check)[0]))
			{
				values.push_back(stof(x->data));  // agr alpha bet nhi ha to number ho ga
			}// number ka alava kuch puch karwana ki zarorat nhi
		}
	}

	bool is_alphabet(char x)
	{
		if (x == '.' || (x >= '1' && x <= '9')) // its not an alphabet
		{
			return false;
		}
		else
			return true; // its  an alphabet
	}

	vector<float> GetInputForApplyingFunction(int& range_start_row, int& range_start_col, int& range_end_col, int& range_end_row)   // this function is for setting the start and end of the range 
	{
		vector<float> values;
		// if single cell is selected
		string val_to_check = this->CellFromWhereRangeForOperationIsStarting->data;
		if (range_start_row == range_end_row && range_end_col == range_start_col)
		{    // agr cell ka data empty nhi ha too 
			value_inserter_for_range_function(this->CellFromWhereRangeForOperationIsStarting, values);
			return values;
		}  // agr row ak hi ho par column different hon (horizontal line case)
		else if (range_end_row == range_start_row)
		{    // agr starting col, ending col (coordinates)  sa bara ha to start or end ko swap kar da
			if (range_start_col > range_end_col)
			{
				// because cell contains pointers will this work ?
				swap(this->CellFromWhereRangeForOperationIsStarting, this->CellFromWhereRangeForOperationIsEnding);
				swap(range_start_col, range_end_col);
			}
			cell* start = CellFromWhereRangeForOperationIsStarting;
			while (start != CellFromWhereRangeForOperationIsEnding)
			{
				value_inserter_for_range_function(start, values);

				start = start->right; // going right
			}
			value_inserter_for_range_function(start, values);
			return values;
		}// same column but change rows
		else if (range_end_col == range_start_col)
		{
			if (range_start_row > range_end_row)
			{
				// because cell contains pointers will this work ?
				swap(this->CellFromWhereRangeForOperationIsStarting, this->CellFromWhereRangeForOperationIsEnding);
				swap(range_start_row, range_end_row);
			}
			cell* start = CellFromWhereRangeForOperationIsStarting;
			while (start != CellFromWhereRangeForOperationIsEnding)
			{
				value_inserter_for_range_function(start, values);
				start = start->down;
			}
			value_inserter_for_range_function(start, values);
			return values;
		}  // both rows and columns are different from each other

		return values;
	}

};
