#include <iostream>
#include <string>
#include <vector>
#include <conio.h>
#include <windows.h>
#include <fstream>
using namespace std;
template <typename T>
class miniExcel
{
public:
    class cell // Cell class
    {
    public:
        T data;
        cell *left;
        cell *right;
        cell *up;
        cell *down;

        cell(T val) // Constructor
        {
            data = val;
            left = nullptr;
            right = nullptr;
            up = nullptr;
            down = nullptr;
        }
    };

    int activeRow = 0; // Atribuutes
    int activeCol = 0;
    int rows;
    int col;
    int boxWidth = 10;
    cell *head;
    cell *active;
    vector<vector<T>> clipboard; // Storing rows and columns for copy and cut
    miniExcel(int rows, int col) : rows(rows), col(col)
    {
        head = nullptr;
        active = nullptr;
        initializeGrid();
    }
    miniExcel()
    {
    }
    class sl_iterator // Iterator class
    {
    private:
        cell *iter;
        friend class miniExcel<T>;

    public:
        sl_iterator(cell *n)
        {
            iter = n;
        }

        sl_iterator begin()
        {
            return sl_iterator(head);
        }

        sl_iterator end()
        {
            return sl_iterator(nullptr);
        }

        sl_iterator operator++()
        {
            if (iter->down != nullptr) // Move down (increment row)
            {
                iter = iter->down;
            }
            return *this;
        }

        sl_iterator operator--()
        {
            if (iter->up != nullptr) // Move up (decrement row)
            {
                iter = iter->up;
            }
            return *this;
        }

        sl_iterator operator++(int)
        {
            if (iter->right != nullptr) // Move right (increment column)
            {
                iter = iter->right;
            }
            return *this;
        }

        sl_iterator operator--(int)
        {
            if (iter->left != nullptr) // Move left (decrement column)
            {
                iter = iter->left;
            }
            return *this;
        }
        T &operator*()
        {
            return iter->data;
        }
        bool operator==(const sl_iterator &other)
        {
            return iter == other.iter;
        }
        bool operator!=(const sl_iterator &other)
        {
            return iter != other.iter;
        }
    };

    void initializeGrid() // Formation of grid
    {
        head = new cell(T{});
        active = head;

        for (int i = 0; i < rows; i++)
        {
            cell *currentRow = active;

            for (int j = 0; j < col; j++)
            {
                cell *newCell = new cell(T{});
                currentRow->right = newCell;
                newCell->left = currentRow;
                currentRow = newCell;

                if (j > 0)
                {
                    newCell->left = currentRow->left;
                    currentRow->left->right = newCell;
                }
            }

            if (i > 0)
            {
                cell *prevRow = active->up;
                prevRow->down = active;
                active->up = prevRow;
            }

            if (i < rows - 1)
            {
                cell *newRow = new cell(T{});
                active->down = newRow;
                newRow->up = active;
                active = newRow;
            }
        }
    }

    T getActiveCellData() // Get cell which is active
    {
        if (active)
        {
            return active->data;
        }
        else
        {
            return 0;
        }
    }

    cell *setActiveCell(int row, int col) // Set active by changing in grid
    {
        if (row >= 0 && row < rows && col >= 0 && col < this->col)
        {
            active = getCell(row, col);
        }
        return active;
    }

    cell *getCell(int row, int col) // Get any cell in the grid
    {
        if (row >= 0 && row < rows && col >= 0 && col < this->col)
        {
            cell *currentCell = head;
            for (int i = 0; i < row; i++)
            {
                if (currentCell)
                {

                    currentCell = currentCell->down;
                }
                else
                {
                    return nullptr;
                }
            }
            for (int j = 0; j < col; j++)
            {
                if (currentCell)
                {
                    currentCell = currentCell->right;
                }
                else
                {
                    return nullptr;
                }
            }
            return currentCell;
        }
        return nullptr;
    }
    void printGrid() // Print Grid
    {
        system("cls");

        int startX = 0;
        int startY = 0;
        for (int i = 0; i < rows; i++)
        {
            int x = startX;
            for (int j = 0; j < col; j++)
            {
                if (i == activeRow && j == activeCol)
                {
                    int k = 14;
                    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE); // Colour
                    SetConsoleTextAttribute(hConsole, k);
                    gotoxy(x, startY);
                    printBox(getActiveCellData(), x, startY);
                    x += boxWidth + 7;
                }
                else
                {
                    int k = 11;
                    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
                    SetConsoleTextAttribute(hConsole, k);
                    gotoxy(x, startY);
                    printBox(getCell(i, j)->data, x, startY);
                    x += boxWidth + 4;
                }
            }
            startY += 3;
        }
    }
    void printBox(T text, int x, int y) // Print box of grid
    {
        string textStr = to_string(text);
        int boxWidth = (textStr.length() * 2) + 4;
        string horizontalLine(boxWidth, '-');
        string emptyLine = "| ";
        string emptyLine2 = " | ";
        gotoxy(x, y);
        cout << horizontalLine << endl;
        gotoxy(x, y + 1);
        cout << emptyLine;
        int padding = (boxWidth - textStr.length() - 2) / 2;
        cout << string(padding, ' ') << textStr << string(padding, ' ');
        cout << emptyLine2 << endl;
        gotoxy(x, y + 2);
        cout << horizontalLine << endl;
    }
    void gotoxy(int x, int y)
    {
        COORD coord;
        coord.X = x;
        coord.Y = y;
        SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coord);
    }
    // Functionalitues to apply
    void insertRowAbove()
    {
        if (active)
        {
            cell *leftmostCell = active;
            while (leftmostCell->left)
            {
                leftmostCell = leftmostCell->left;
            }

            cell *newRow = new cell(T{});
            cell *currentCell = newRow;

            for (int i = 1; i < col; i++)
            {
                cell *newCell = new cell(T{});
                currentCell->right = newCell;
                newCell->left = currentCell;
                currentCell = newCell;
            }
            if (leftmostCell->up == nullptr)
            {
                newRow->up = nullptr;
                newRow->down = leftmostCell;
                leftmostCell->up = newRow;
                head = newRow;
            }
            else
            {
                newRow->up = leftmostCell->up;
                leftmostCell->up->down = newRow;
                leftmostCell->up = newRow;
                newRow->down = leftmostCell;
            }

            rows++;
        }
    }
    void insertRowBelow()
    {
        if (active)
        {
            cell *leftmostCell = active;
            while (leftmostCell->left)
            {
                leftmostCell = leftmostCell->left;
            }

            cell *newRow = new cell(T{});
            cell *currentCell = newRow;

            for (int i = 1; i < col; i++)
            {
                cell *newCell = new cell(T{});
                currentCell->right = newCell;
                newCell->left = currentCell;
                currentCell = newCell;
            }
            if (leftmostCell->down == nullptr)
            {
                newRow->down = nullptr;
                newRow->up = leftmostCell;
                leftmostCell->down = newRow;
            }
            else
            {
                newRow->down = leftmostCell->down;
                leftmostCell->down->up = newRow;
                leftmostCell->down = newRow;
                newRow->up = leftmostCell;
            }
            rows++;
        }
    }
    void insertColumnRight()
    {
        if (active)
        {
            cell *topmostCell = active;
            while (topmostCell->up)
            {
                topmostCell = topmostCell->up;
            }
            cell *newColumn = new cell(T{});
            cell *current = newColumn;

            while (topmostCell)
            {
                if (topmostCell->right == nullptr)
                {
                    topmostCell->right = newColumn;
                    newColumn->left = topmostCell;
                }
                else
                {

                    cell *rightCell = topmostCell->right;
                    topmostCell->right = newColumn;
                    newColumn->left = topmostCell;
                    newColumn->right = rightCell;
                    rightCell->left = newColumn;
                }
                topmostCell = topmostCell->down;

                if (topmostCell)
                {
                    cell *nextRow = new cell(T{});
                    current->down = nextRow;
                    nextRow->up = current;
                    current = nextRow;
                }
            }
            col++;
            activeCol++;
        }
    }
    void DeleteColumn()
    {
        if (active && col > 1)
        {
            cell *currentCol = active;
            while (currentCol->up != nullptr)
            {
                currentCol = currentCol->up;
            }
            cell *colLeft = currentCol->left;
            cell *colRight = currentCol->right;
            if (colLeft)
            {
                colLeft->right = colRight;
            }
            if (colRight)
            {
                colRight->left = colLeft;
            }
            if (currentCol == head)
            {
                head = colRight;
            }
            active = colRight;
            deleteColCells(currentCol);
            col--;
        }
    }

    void deleteColCells(cell *row)
    {
        cell *currentCell = row;
        while (currentCell)
        {
            cell *nextCell = currentCell->down;
            delete currentCell;
            currentCell = nextCell;
        }
    }
    void insertColumnLeft()
    {
        if (active)
        {
            for (int i = 0; i < rows; i++)
            {
                cell *currentCell = getCell(i, activeCol);
                cell *newCell = new cell(0);
                newCell->right = currentCell;
                if (activeCol > 0)
                {
                    cell *leftCell = getCell(i, activeCol - 1);
                    leftCell->right = newCell;
                    newCell->left = leftCell;
                }
                else if (activeCol == 0 && i == 0)
                {
                    head = newCell;
                }
                else if (activeCol == 0)
                {
                    newCell->left = nullptr;
                }
            }
            col++;
            activeCol++;

            if (activeCol < col)
            {
                active = getCell(activeRow, activeCol);
            }
        }
    }
    void DeleteRow()
    {
        if (active && rows > 1)
        {
            cell *currentRow = active;
            while (currentRow->left != nullptr)
            {
                currentRow = currentRow->left;
            }
            cell *rowAbove = currentRow->up;
            cell *rowBelow = currentRow->down;
            if (rowAbove)

            {
                rowAbove->down = rowBelow;
            }
            if (rowBelow)
            {
                rowBelow->up = rowAbove;
            }
            if (currentRow == head)
            {
                head = rowBelow;
            }
            deleteRowCells(currentRow);
            rows--;
            active = rowBelow;
        }
    }
    void deleteRowCells(cell *row)
    {
        cell *currentCell = row;
        while (currentCell)
        {
            cell *nextCell = currentCell->right;
            delete currentCell;
            currentCell = nextCell;
        }
    }
    void setInputToActiveCell()
    {
        if (active)
        {
            gotoxy(activeCol * (boxWidth + 2) + 2, activeRow * 4 + 1);
            cout << "| ";
            cin >> active->data;
            printGrid();
            gotoxy(activeCol * (boxWidth + 2) + 2, activeRow * 4 + 1);
            cout << "| ";
        }
    }
    void clearRowData()
    {
        cell *temp = active;

        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        while (temp != nullptr)
        {
            cell *newnode = new cell(T());
            temp->data = newnode->data;
            temp = temp->right;
        }
    }
    void ClearColumn()
    {
        for (int i = 0; i < rows; ++i)
        {
            cell *currentCell = getCell(i, activeCol);
            if (currentCell)
            {
                currentCell->data = 0;
            }
        }
    }
    void RangedSum()
    {
        cout << "Finding sum of value from given range" << endl;
        int startRow, startCol, endRow, endCol;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int sum = GetRangeSum(startRow, startCol, endRow, endCol);
        cout << "Sum of the specified range: " << sum << endl;
    }
    T GetRangeSum(int startRow, int startCol, int endRow, int endCol)
    {
        T sum = 0;
        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
            return 0;
        }
        else
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        sum += currentCell->data;
                    }
                }
            }
        }
        return sum;
    }
    void RangedAverage()
    {
        cout << "Finding average value from given range" << endl;
        int startRow, startCol, endRow, endCol;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int average = GetRangeAverage(startRow, startCol, endRow, endCol);
        cout << "Average of the specified range: " << average << endl;
    }
    T GetRangeAverage(int startRow, int startCol, int endRow, int endCol)
    {

        T sum = 0;
        T count = 0;
        T average = 0;
        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
            return 0;
        }
        else
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        sum += currentCell->data;
                        count++;
                    }
                }
            }
        }
        if (count > 0)
        {
            average = sum / count;
        }
        return average;
    }
    T GetRangeCount(int startRow, int startCol, int endRow, int endCol)
    {

        T count = 0;
        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
            return 0;
        }
        else
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        count++;
                    }
                }
            }
        }
        return count;
    }
    void RangeCount()
    {
        cout << "Finding number of cell  from given range" << endl;
        int startRow, startCol, endRow, endCol, row, column;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int count = GetRangeCount(startRow, startCol, endRow, endCol);
        cout << "Select cell in which you want to print sum" << endl;
        cout << "Enter row number";
        cin >> row;
        cout << "Enter column number ";
        cin >> column;
        cell *current = getCell(row, column);
        current->data = count;
    }
    T GetMaxValue(int startRow, int startCol, int endRow, int endCol)
    {

        T maxValue = 0;
        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
            return 0;
        }
        else
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        if (currentCell->data > maxValue)
                        {
                            maxValue = currentCell->data;
                        }
                    }
                }
            }
        }
        return maxValue;
    }
    void RangeMax()
    {
        cout << "Finding maximum value from given range" << endl;
        int startRow, startCol, endRow, endCol, row, column;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int max = GetMaxValue(startRow, startCol, endRow, endCol);
        cout << "Select cell in which you want to print sum" << endl;
        cout << "Enter row number";
        cin >> row;
        cout << "Enter column number ";
        cin >> column;
        cell *current = getCell(row, column);
        current->data = max;
    }
    T GetMinValue(int startRow, int startCol, int endRow, int endCol)
    {

        T minValue = 100000;
        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
            return 0;
        }
        else
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        if (currentCell->data < minValue)
                        {
                            minValue = currentCell->data;
                        }
                    }
                }
            }
        }
        return minValue;
    }
    void RangeMin()
    {
        cout << "Finding minimum value from given range" << endl;
        int startRow, startCol, endRow, endCol, row, column;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int min = GetMinValue(startRow, startCol, endRow, endCol);
        cout << "Select cell in which you want to print sum" << endl;
        cout << "Enter row number";
        cin >> row;
        cout << "Enter column number ";
        cin >> column;
        cell *current = getCell(row, column);
        current->data = min;
    }
    void RangeSum()
    {
        cout << "Finding sumn value from given range and display in grid " << endl;
        int startRow, startCol, endRow, endCol, row, column;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int sum = GetRangeSum(startRow, startCol, endRow, endCol);
        cout << "Select cell in which you want to print sum" << endl;
        cout << "Enter row number";
        cin >> row;
        cout << "Enter column number ";
        cin >> column;
        cell *current = getCell(row, column);
        current->data = sum;
    }
    void RangeAverage()
    {
        cout << "Finding average value from given range and display in grid " << endl;
        int startRow, startCol, endRow, endCol, row, column;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        int average = GetRangeAverage(startRow, startCol, endRow, endCol);
        cout << "Select cell in which you want to print sum" << endl;
        cout << "Enter row number";
        cin >> row;
        cout << "Enter column number ";
        cin >> column;
        cell *current = getCell(row, column);
        current->data = average;
    }
    void RangeCopy()
    {
        cout << "Copy data from grid" << endl;
        int startRow, startCol, endRow, endCol;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        copy(startRow, startCol, endRow, endCol);
    }
    void copy(int startRow, int startCol, int endRow, int endCol)
    {

        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
        }
        else
        {
            clipboard.clear();

            for (int row = startRow; row <= endRow; row++)
            {
                vector<T> rowData;
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        rowData.push_back(currentCell->data);
                    }
                }
                clipboard.push_back(rowData);
            }
        }
        cout << "Data  is successfully copy  and save in the grid";
    }
    void cut(int startRow, int startCol, int endRow, int endCol)
    {

        if (startRow >= endRow || startCol >= endCol)
        {
            cout << "Invalid Selection" << endl;
        }
        else
        {
            clipboard.clear(); // Clear the clipboard before copying new data

            for (int row = startRow; row <= endRow; row++)
            {
                vector<T> rowData;
                for (int col = startCol; col <= endCol; col++)
                {
                    cell *currentCell = getCell(row, col);
                    if (currentCell)
                    {
                        rowData.push_back(currentCell->data);
                        currentCell->data = T();
                    }
                }
                clipboard.push_back(rowData);
            }
        }
        cout << "Data  is successfully cut and save in vector";
    }
    void RangeCut()
    {
        cout << "Cut data from grid" << endl;
        int startRow, startCol, endRow, endCol;
        cout << "Enter start row: ";
        cin >> startRow;
        cout << "Enter start col: ";
        cin >> startCol;
        cout << "Enter end row: ";
        cin >> endRow;
        cout << "Enter end col: ";
        cin >> endCol;
        cut(startRow, startCol, endRow, endCol);
    }
    void Paste()
    {
        cout << "Paste data from clipboard" << endl;
        if (clipboard.empty() || clipboard[0].empty())
        {
            cout << "Clipboard is empty." << endl;
            return;
        }
        int clipboardRows = clipboard.size();
        int clipboardCols = clipboard[0].size();
        int targetRow = activeRow;
        int targetCol = activeCol;
        for (int i = 0; i < clipboardRows; i++)
        {
            for (int j = 0; j < clipboardCols; j++)
            {
                cell *targetCell = getCell(targetRow + i, targetCol + j);
                if (targetCell)
                {
                    targetCell->data = clipboard[i][j];
                }
            }
        }
        cout << "Data Paste in the grid" << endl;
    }
    void storeGrid()
    {
        fstream file;
        file.open("grid.txt", ios::out);
        if (file.is_open())
        {
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    T cellData = getCell(i, j)->data;
                    file << cellData << "\t";
                }
                file << endl;
            }
            file.close();
            cout << "Grid data stored in 'grid.txt'." << endl;
        }
        else
        {
            cout << "Unable to open the file for storing grid data." << endl;
        }
    }
    void loadData(const string &filename)
    {
        ifstream file(filename);
        if (file.is_open())
        {
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    T cellData;
                    file >> cellData;
                    getCell(i, j)->data = cellData;
                }
            }
            file.close();
            cout << "Grid data loaded from '" << filename << "'." << endl;
        }
        else
        {
            cout << "Unable to open the file for loading grid data." << endl;
        }
    }
    void incrementActiveRow()
    {
        sl_iterator iter(active);
        iter++;
        if (iter != sl_iterator(nullptr))
        {
            active = iter.iter;
            activeRow++;
        }
    }
    void decrementActiveRow()
    {
        sl_iterator iter(active);
        iter--;
        if (iter != sl_iterator(nullptr))
        {
            active = iter.iter;
            activeRow--;
        }
    }
    void incrementActiveColumn()
    {
        sl_iterator iter(active);
        iter++;
        if (iter != sl_iterator(nullptr))
        {
            active = iter.iter;
            activeCol++;
        }
    }
    void decrementActiveColumn()
    {
        sl_iterator iter(active);
        iter--;
        if (iter != sl_iterator(nullptr))
        {
            active = iter.iter;
            activeCol--;
        }
    }
    void InsertCellByDownShift()
    {
        if (active)
        {
            cell *belowCell = active->down;
            if (belowCell)
            {
                cell *newCell = new cell(0);
                active->data = newCell->data;
                active = belowCell;
            }
            else
            {
                insertRowBelow();
                active = active->down;
            }
            active = belowCell;
            activeRow++;
        }
    }
    void InsertCellByRightShift()
    {
        if (active)
        {
            cell *rightCell = active->right;
            if (rightCell)
            {
                cell *newCell = new cell(0);
                active->data = newCell->data;
                active = rightCell;
            }
            else
            {
                insertColumnRight();
                active = active->right;
            }
            active = rightCell;
            activeCol++;
        }
    }
    void DeleteCellByLeftShift()
    {
        if (active && active->left != nullptr)
        {
            cell *nextActive = active->left;
            nextActive->right = active->right;
            if (active->right != nullptr)
            {
                active->right->left = nextActive;
            }
            delete active;
            active = nextActive;
            activeCol--;
        }
        cell *currentCell = active;
        while (currentCell->right != nullptr)
        {
            currentCell->data = currentCell->right->data;
            currentCell = currentCell->right;
        }
        currentCell->data = 0;
    }
    void DeleteCellByUpShift()
    {
        if (active && active->up != nullptr)
        {
            cell *nextActive = active->up;
            nextActive->down = active->down;
            if (active->down != nullptr)
            {
                active->down->up = nextActive;
            }
            delete active;
            active = nextActive;
            activeRow--;
        }
        cell *currentCell = active;
        while (currentCell->down != nullptr)
        {
            currentCell->data = currentCell->down->data;
            currentCell = currentCell->down;
        }
        currentCell->data = 0;
    }
};
void header()
{
    miniExcel<int> m1;
    int k = 11;
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    SetConsoleTextAttribute(hConsole, k);
    m1.gotoxy(40, 2);
    cout << " __  __   _           _     _____                        _ ";
    m1.gotoxy(40, 3);
    cout << "|  \\/  | (_)  _ __   (_)   | ____| __  __   ___    ___  | | ";
    m1.gotoxy(40, 4);
    cout << "| |\\/| | | | | '_ \\  | |   |  _|   \\ \\/ /  / __|  / _ \\ | |";
    m1.gotoxy(40, 5);
    cout << "| |  | | | | | | | | | |   | |___   >  <  | (__  |  __/ | | ";
    m1.gotoxy(40, 6);
    cout << "|_|  |_| |_| |_| |_| |_|   |_____| /_/\\_\\  \\___|  \\___| |_|";
    m1.gotoxy(40, 7);
    SetConsoleTextAttribute(hConsole, 14);
}
void Menu()
{
    miniExcel<int> m1;
    m1.gotoxy(2, 10);
    cout << "-------------------------------------";
    m1.gotoxy(2, 11);
    cout << "Functions you can apply";
    m1.gotoxy(2, 12);
    cout << "---------------------------------------";
    m1.gotoxy(2, 13);
}
void Functionalities()
{
    Menu();
    miniExcel<int> m1;
    m1.gotoxy(2, 14);
    cout << "1.Print sheet";
    m1.gotoxy(2, 15);
    cout << "2.Row operations";
    m1.gotoxy(2, 16);
    cout << "3.Column Operations";
    m1.gotoxy(2, 17);
    cout << "4.Cell Operations";
    m1.gotoxy(2, 18);
    cout << "5.Calcualtions";
    m1.gotoxy(2, 19);
    cout << "6.Edit";
    m1.gotoxy(2, 20);
    cout << "7.File Handling";
    m1.gotoxy(2, 21);
    cout << "8.Exit";
    m1.gotoxy(2, 22);
}
string showMenu()
{
    miniExcel<int> m1;
    Functionalities();
    string choice;
    system("cls");
    header();
    Functionalities();
    cout << "Enter your choice ";
    m1.gotoxy(2, 23);
    cin >> choice;
    return choice;
}
void showScreen()
{
    int initialRows = 5;
    int initialCols = 5;
    miniExcel<int> m1(initialRows, initialCols);
    m1.setActiveCell(m1.activeRow, m1.activeCol);
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    m1.loadData("grid.txt");
    string choice = showMenu();
    header();
    system("cls");
    while (choice != "8")
    {
        if (choice == "1")
        {
            m1.printGrid();
            m1.gotoxy(75, 12);
            cout << "1.Press '*' to go back";
        }
        if (choice == "2")
        {
            m1.printGrid();
            m1.gotoxy(75, 12);
            cout << "1.Press 'A' or 'a' to insert a row above";
            m1.gotoxy(75, 13);
            cout << "2.Press 'B' or 'b' to insert a row above";
            m1.gotoxy(75, 14);
            cout << "3.Press 'W' or 'w' to delete a row ";
            m1.gotoxy(75, 15);
            cout << "4.Press 'K' or 'k' to clear a row";
            m1.gotoxy(75, 16);
            cout << "5.Press '*' to go back";
            m1.gotoxy(75, 17);
        }
        if (choice == "3")
        {

            m1.printGrid();
            m1.gotoxy(75, 21);
            cout << "1.Press 'E' or 'e' to insert a column right";
            m1.gotoxy(75, 22);
            cout << "2.Press 'F' or 'f' to insert a column left";
            m1.gotoxy(75, 23);
            cout << "3.Press 'G' or 'g' to delete a column ";
            m1.gotoxy(75, 24);
            cout << "4.Press 'H' or 'h' to clear a column";
            m1.gotoxy(75, 25);
            cout << "5.Press '*'  to go back";
            m1.gotoxy(75, 26);
        }
        else if (choice == "4")
        {
            m1.printGrid();
            m1.gotoxy(75, 21);
            cout << "1.Press 'R' or 'r' to insert cell by right shift";
            m1.gotoxy(75, 22);
            cout << "2.Press 'd' or 'D' to insert a cell by down shift";
            m1.gotoxy(75, 23);
            cout << "3.Press 'l' or 'l' to delete a cell by left shift ";
            m1.gotoxy(75, 24);
            cout << "4.Press 'u' or 'U' to delete a cell by up shift";
            m1.gotoxy(75, 25);
            cout << "5.Press '*' to go back";
            m1.gotoxy(75, 26);
        }
        else if (choice == "5")
        {
            m1.printGrid();
            m1.gotoxy(75, 21);
            cout << "1.Press 'S' or 's' to find sum ";
            m1.gotoxy(75, 22);
            cout << "2.Press 'j' or 'J' to find average ";
            m1.gotoxy(75, 23);
            cout << "3.Press 'M' or 'm' to  find max and dispaly in grid";
            m1.gotoxy(75, 24);
            cout << "4.Press 'N' or 'n' to  find min and dispaly in grid";
            m1.gotoxy(75, 25);
            cout << "5.Press 'O' or 'o' to  find count and dispaly in grid";
            m1.gotoxy(75, 26);
            cout << "6.Press 'i' or 'i' to  find sum and dispaly in grid";
            m1.gotoxy(75, 27);
            cout << "6.Press 'y' or 'y' to  find average and dispaly in grid";
            m1.gotoxy(75, 28);
            cout << "7.Press '*'  to go back";
            m1.gotoxy(75, 29);
        }
        else if (choice == "6")
        {
            m1.printGrid();
            m1.gotoxy(75, 21);
            cout << "1.Press 'C' or 'c' to copy";
            m1.gotoxy(75, 22);
            cout << "2.Press 'X' or 'x' to  cut";
            m1.gotoxy(75, 23);
            cout << "3.Press 'V' or 'v' to  paste";
            m1.gotoxy(75, 24);
            cout << "4.Press '*' or '*' to go back";
            m1.gotoxy(75, 25);
        }
        else if (choice == "7")
        {
            m1.printGrid();
            m1.gotoxy(75, 12);
            cout << "Press 'P' or 'p' to store";
            m1.gotoxy(75, 13);
            cout << "Press '*' to go back";
            m1.gotoxy(75, 13);
        }

        int key = _getch();
        if (key == 224)
        {
            int arrowKey = _getch();
            if (arrowKey == 72 && m1.activeRow > 0) // UP Arrow Key
            {
                m1.decrementActiveRow();
            }
            else if (arrowKey == 80 && m1.activeRow < m1.rows - 1) // Down Arrow
            {
                m1.incrementActiveRow();
            }
            else if (arrowKey == 75 && m1.activeCol > 0) // Left Arrow
            {
                m1.decrementActiveColumn();
            }
            else if (arrowKey == 77 && m1.activeCol < m1.col - 1) // Riight Arrow
            {
                m1.incrementActiveColumn();
            }
            m1.setActiveCell(m1.activeRow, m1.activeCol);
        }
        else if (key == 65 || key == 97 && choice == "2") // A OR a
        {
            m1.insertRowAbove();
        }
        else if (key == 66 || key == 98 && choice == "2") // B or b
        {
            m1.insertRowBelow();
        }
        else if (key == 107 || key == 75 && choice == "2") // K or k
        {
            m1.clearRowData();
        }
        else if (key == 119 || key == 87 && choice == "2") // W or w
        {
            m1.DeleteRow();
        }
        else if (key == 69 || key == 101 && choice == "3") // E or e
        {
            m1.insertColumnRight();
        }
        else if (key == 70 || key == 102 && choice == "3") // F or f
        {
            m1.insertColumnLeft();
        }
        else if (key == 71 || key == 103 && choice == "3") // G or g
        {
            m1.DeleteColumn();
        }
        else if (key == 72 || key == 104 && choice == "3") // H or h
        {
            m1.ClearColumn();
        }
        else if (key == 42) //*
        {
            system("cls");
            choice = showMenu();
        }
        else if (key == 13) // enter
        {
            m1.setInputToActiveCell();
        }
        else if (key == 83 || key == 115 && choice == "5") // S or s
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangedSum();
            getch();
        }
        else if (key == 74 || key == 106 && choice == "5") // J or j
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangedAverage();
            getch();
        }
        else if (key == 79 || key == 111 && choice == "5") // O or o
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeCount();
            getch();
        }
        else if (key == 77 || key == 109 && choice == "5") // M or m
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeMax();
            getch();
        }
        else if (key == 78 || key == 110 && choice == "5") // N or n
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeMin();
            getch();
        }
        else if (key == 105 || key == 73 && choice == "5") // I or i
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeSum();
            getch();
        }
        else if (key == 89 || key == 121 && choice == "5") // Y or y
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeAverage();
            getch();
        }
        else if (key == 67 || key == 99 && choice == "6") // C or c
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeCopy();
            getch();
        }
        else if (key == 88 || key == 120 && choice == "6") // X or x
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.RangeCut();
            getch();
        }
        else if (key == 118 || key == 86 && choice == "6") // V  or v
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.Paste();
            getch();
        }
        else if (key == 80 || key == 112 && choice == "7") // P or p
        {
            m1.gotoxy(1, 18);
            SetConsoleTextAttribute(hConsole, 14);
            m1.storeGrid();
        }
        else if (key == 100 || key == 68 && choice == "4") // D or d
        {
            m1.InsertCellByDownShift();
        }
        else if (key == 114 || key == 82 && choice == "4") // R or r
        {
            m1.InsertCellByRightShift();
        }
        else if (key == 76 || key == 108 && choice == "4") // L or L
        {
            m1.DeleteCellByLeftShift();
        }
        else if (key == 85 || key == 117 && choice == "4") // u or U
        {
            m1.DeleteCellByUpShift();
        }
    }
}
int main()
{
    system("cls");
    Sleep(100);
    header();
    showScreen();
    return 0;
}