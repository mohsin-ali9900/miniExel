#include <iostream>
#include <windows.h>
#include <string>
#include <vector>
#include <conio.h>
#include <exception>
#include <limits>
using namespace std;
void hideCursor();
void ClearScreen();
template <typename T>
class Node
{
public:
    Node<T>* left;
    Node<T>* right;
    Node<T>* up;
    Node<T>* down;
    T data;
    Node(T val)
    {
        left = nullptr;
        right = nullptr;
        up = nullptr;
        down = nullptr;
        data = val;
    }
};

class Coordinate
{
public:
    int row;
    int col;
    Coordinate(int row, int col)
    {
        this->row = row;
        this->col = col;
    }
};
template <typename T>
class MiniExcel
{
private:
    Node<T>* root;
    Node<T>* current;
    int rows;
    int cols;

    void SetConsoleColor(int color)
    {
        SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
    }

public:
    MiniExcel(int row, int col)
    {
        root = nullptr;
        current = nullptr;
        rows = row;
        cols = col;
        CreateDefaultGrid();
    }

    void CreateDefaultGrid()
    {
        root = new Node<T>("     ");
        Node<T>* prevRowStart = root;
        Node<T>* prevColStart = root;
        for (int col = 1; col < cols; col++)
        {
            Node<T>* newcol = new Node<T>("     ");
            newcol->left = prevColStart;
            prevColStart->right = newcol;
            prevColStart = newcol;
        }

        for (int row = 1; row < rows; row++)
        {
            Node<T>* newRow = new Node<T>("     ");
            newRow->up = prevRowStart;
            prevRowStart->down = newRow;
            prevColStart = newRow;
            for (int col = 1; col < cols; col++)
            {
                Node<T>* newcol = new Node<T>("     ");
                if (prevColStart->up != nullptr)
                {
                    newcol->up = prevColStart->up->right;
                    prevColStart->up->right->down = newcol;
                }
                newcol->left = prevColStart;
                prevColStart->right = newcol;
                prevColStart = newcol;
            }
            prevRowStart = newRow;
        }
        current = root;
    }

    void HighlightCurrentCell()
    {
        SetConsoleColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY);
    }

    void PrintGrid()
    {
        char c = 61;
        HANDLE H = GetStdHandle(STD_OUTPUT_HANDLE);
        Node<T>* temp = root;
        while (temp != nullptr)
        {
            Node<T>* currentNode = temp;
            while (currentNode != nullptr)
            {
                cout << "+" << c << c << c << c << c << c;
                currentNode = currentNode->right;
            }
            cout << "+" << endl;

            currentNode = temp;
            while (currentNode != nullptr)
            {
                if (currentNode == current)
                {

                    SetConsoleTextAttribute(H, 2);
                    cout << "|";
                    SetConsoleTextAttribute(H, 2);
                    if (currentNode->data.length() < 4)
                    {
                        for (int i = currentNode->data.length(); i <= 4; i++)
                        {
                            cout << " ";
                        }
                        cout << currentNode->data << " ";
                    }
                    else if (current->data.length() == 4)
                    {
                        cout << " ";
                        cout << currentNode->data << " ";
                    }
                    else
                    {
                        string sub = currentNode->data.substr(0, 4);
                        cout << " " << sub << " ";
                    }
                    SetConsoleTextAttribute(H, 7);
                }
                else
                {
                    cout << "|";
                    if (currentNode->data.length() < 4)
                    {
                        for (int i = currentNode->data.length(); i <= 4; i++)
                        {
                            cout << " ";
                        }
                        cout << currentNode->data << " ";
                    }
                    else if (this->current->data.length() == 4)
                    {
                        cout << " ";
                        cout << currentNode->data << " ";
                    }
                    else
                    {
                        string sub = currentNode->data.substr(0, 4);
                        cout << " " << sub << " ";
                    }
                }
                currentNode = currentNode->right;
            }
            cout << "|" << endl;
            temp = temp->down;
        }
        Node<T>* current = root;
        while (current != nullptr)
        {
            cout << "+" << c << c << c << c << c << c;
            current = current->right;
        }
        cout << "+" << endl;
    }
    void InsertColumnAtRight()
    {
        Node<T>* node = root;
        while (node->right != nullptr)
        {
            node = node->right;
        }

        while (node != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            node->right = newNode;
            newNode->left = node;
            if (node->up != nullptr)
            {
                newNode->up = node->up->right;
                if (node->up->right != nullptr)
                {
                    node->up->right->down = newNode;
                }
            }
            node = node->down;
        }
        cols++;
    }

    void InsertRowAtEnd()
    {
        Node<T>* downNode = root;
        while (downNode->down != nullptr)
        {
            downNode = downNode->down;
        }

        while (downNode != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            downNode->down = newNode;
            newNode->up = downNode;
            if (downNode->left != nullptr)
            {
                newNode->left = downNode->left->down;
                if (downNode->left->down != nullptr)
                {
                    downNode->left->down->right = newNode;
                }
            }
            downNode = downNode->right;
        }
        rows++;
    }
    void InsertCellByRightShift()
    {
        InsertColumnAtRight();
        Node<T>* last = current;
        while (last->right != nullptr)
        {
            last = last->right;
        }
        Node<T>* chase = last;
        while (chase != current)
        {
            chase->data = chase->left->data;
            chase = chase->left;
        }
        current->data = " ";
        cols++;
    }

    void InsertCellByDownShift()
    {
        InsertRowAtEnd();
        Node<T>* last = current;
        while (last->down != nullptr)
        {
            last = last->down;
        }
        Node<T>* chase = last;
        while (chase != current)
        {
            chase->data = chase->up->data;
            chase = chase->up;
        }
        current->data = " ";
        rows++;
    }

    void DeleteCellByLeftShift()
    {
        if (current->right == nullptr && current->left == nullptr && current->down != nullptr)
        {
            Node<T>* upper = current->up;
            Node<T>* lower = current->down;
            while (upper != nullptr)
            {
                upper->down = lower;
                lower->up = upper;
                upper = upper->right;
                lower = lower->right;
            }
        }
        else if (current->right == nullptr)
        {
            current = nullptr;
        }
        else
        {
            Node<T>* temp = current;
            while (temp->right->right != nullptr)
            {
                temp->data = temp->right->data;
                temp = temp->right;
            }
            if (temp->right->down != nullptr and temp->right->up != nullptr)
            {
                temp->right->down->up = nullptr;
                temp->right->up->down = nullptr;
            }

            temp->right = nullptr;
        }
        if (current->left != nullptr)
        {
            current = current->left;
        }
    }

    void DeleteCellByUpShift()
    {
        if (current->up == nullptr && current->down == nullptr && current->right != nullptr)
        {
            Node<T>* lef = current->left;
            Node<T>* righ = current->right;
            while (lef != nullptr)
            {
                lef->right = righ;
                righ->left = lef;
                righ = righ->down;
                lef = lef->down;
            }
        }
        else if (current->down == nullptr)
        {
            current = nullptr;
        }
        else
        {
            Node<T>* temp = current;
            while (temp->down->down != nullptr)
            {
                temp->data = temp->down->data;
                temp = temp->down;
            }
            if (temp->down->left != nullptr && temp->down->right != nullptr)
            {
                temp->down->left->right = nullptr;
                temp->down->right->left = nullptr;
            }
            temp->down = nullptr;
        }
    }

    void SetCurrentToLeft()
    {
        current = current->left;
    }
    void SetCurrentToRight()
    {
        current = current->right;
    }
    void SetCurrentToUp()
    {
        current = current->up;
    }
    void SetCurrentToDown()
    {
        current = current->down;
    }

    Node<T>* GetCurrentCell()
    {
        return current;
    }

    void SetCurrentData(T val)
    {
        current->data = val;
    }

    void TakeInput()
    {
        string input;
        getline(cin, input);
        SetCurrentData(input);
        input = "";
    }

    void InserRowAbove()
    {
        Node<T>* node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }

        while (node != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            newNode->up = node->up;
            if (node->up != nullptr)
            {
                node->up->down = newNode;
            }
            newNode->down = node;
            node->up = newNode;
            if (node->left != nullptr)
            {
                newNode->left = node->left->up;
                if (node->left->up != nullptr)
                {
                    node->left->up->right = newNode;
                }
            }
            node = node->right;
        }
        if (root->up != nullptr)
        {
            root = root->up;
        }
        rows++;
    }

    void InserRowBelow()
    {
        Node<T>* downNode = current;
        while (downNode->left != nullptr)
        {
            downNode = downNode->left;
        }

        while (downNode != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            newNode->down = downNode->down;
            if (downNode->down != nullptr)
            {
                downNode->down->up = newNode;
            }
            newNode->up = downNode;
            downNode->down = newNode;
            if (downNode->left != nullptr)
            {
                newNode->left = downNode->left->down;
                if (downNode->left->down != nullptr)
                {
                    downNode->left->down->right = newNode;
                }
            }
            downNode = downNode->right;
        }
        rows++;
    }

    void InsertColumnToLeft()
    {
        Node<T>* node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }

        while (node != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            newNode->left = node->left;
            if (node->left != nullptr)
            {
                node->left->right = newNode;
            }
            newNode->right = node;
            node->left = newNode;
            if (node->up != nullptr)
            {
                newNode->up = node->up->left;
                if (node->up->left != nullptr)
                {
                    node->left->up->down = newNode;
                }
            }
            node = node->down;
        }
        // if (current == root)
        if (root->left != nullptr)
        {
            root = root->left;
        }
        cols++;
    }

    void InsertColumnToRight()
    {
        Node<T>* node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }
        while (node != nullptr)
        {
            Node<T>* newNode = new Node<T>("   ");
            newNode->right = node->right;
            if (node->right != nullptr)
            {
                node->right->left = newNode;
            }
            newNode->left = node;
            node->right = newNode;
            if (node->up != nullptr)
            {
                newNode->up = node->up->right;
                if (node->up->right != nullptr)
                {
                    node->up->right->down = newNode;
                }
            }
            node = node->down;
        }
        cols++;
    }

    void DeleteRow()
    {
        // Node<T> *node = current;
        // while (node->left != nullptr)
        // {
        //     node = node->left;
        // }

        // while (node != nullptr)
        // {
        //     if (node->up == nullptr)
        //     {
        //         node->down->up = nullptr;
        //     }
        //     else if (node->down == nullptr)
        //     {
        //         node->up->down = nullptr;
        //     }
        //     else
        //     {
        //         node->down->up = node->up;
        //         node->up->down = node->down;
        //     }
        //     node = node->right;
        // }
        // if (current->up == nullptr)
        // {
        //     root = root->down;
        //     current = current->down;
        // }
        // else if (current->down == nullptr)
        // {
        //     current = current->up;
        // }
        // else
        // {
        //     current = current->down;
        // }
        // rows--;
        Node<T>* node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }

        while (node != nullptr)
        {
            if (node->up == nullptr)
            {
                node->down->up = nullptr;
            }
            else if (node->down == nullptr)
            {
                node->up->down = nullptr;
            }
            else
            {
                node->down->up = node->up;
                node->up->down = node->down;
            }
            node = node->right;
        }
        if (current->up == nullptr)
        {
            root = root->down;
            current = current->down;
        }
        else if (current->down == nullptr)
        {
            current = current->up;
        }
        else
        {
            current = current->down;
        }
        rows--;
    }

    void DeleteColumn()
    {
        Node<T>* node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }

        while (node != nullptr)
        {
            if (node->right == nullptr)
            {
                node->left->right = nullptr;
            }
            else if (node->left == nullptr)
            {
                node->right->left = nullptr;
            }
            else
            {
                node->left->right = node->right;
                node->right->left = node->left;
            }
            node = node->down;
        }
        if (current->left == nullptr)
        {
            root = root->right;
            current = current->right;
        }
        else if (current->down == nullptr)
        {
            current = current->left;
        }
        else
        {
            current = current->left;
        }
        cols--;
    }

    void ClearColumn()
    {
        Node<T>* node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }
        while (node != nullptr)
        {
            node->data = "     ";
            node = node->down;
        }
    }

    void ClearRow()
    {
        Node<T>* node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }
        while (node != nullptr)
        {
            node->data = "     ";
            node = node->right;
        }
    }

    int getCurrentData(string str)
    {
        int data = 0;
        string tempString = "";
        for (char c : str)
        {
            if (c >= '0' && c <= '9')
            {
                tempString += c;
            }
        }
        if (tempString != "")
        {
            data = stoi(tempString);
        }
        return data;
    }

    int GetRangeSum(int startRow, int startColumn, int endRow, int endColumn)
    {
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            return 0;
        }

        int sum = 0;
        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = getCurrentData(currentCol->data);
                sum += data;
            }
        }

        return sum;
    }

    int GetRangeAverge(int startRow, int startColumn, int endRow, int endColumn)
    {
        int count = 0;
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            return 0;
        }

        int sum = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = getCurrentData(currentCol->data);
                sum += data;
                count++;
            }
        }

        return sum / count;
    }

    // void Sum(int startRow, int startColumn, int endRow, int endColumn)
    // {
    //     int sum = GetRangeSum(startRow, startColumn, endRow, endColumn);
    //     current->data = sum;
    // }

    // void Average(int startRow, int startColumn, int endRow, int endColumn)
    // {
    //     int avg = GetRangeAverge(startRow, startColumn, endRow, endColumn);
    //     current->data = avg;
    // }

    void Count(int startRow, int startColumn, int endRow, int endColumn)
    {
        int count = 0;
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            count = 0;
        }

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                count++;
            }
        }

        current->data = to_string(count);
    }

    void Minimum(int startRow, int startColumn, int endRow, int endColumn)
    {
        // int min = numeric_limits<int>::min();
        int min = 2147483647;
        int minValue = min;
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            min = minValue;
        }

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = getCurrentData(currentCol->data);
                if (data < min)
                {
                    min = data;
                }
            }
        }
        if (min != minValue)
        {
            current->data = to_string(min);
        }
    }
    void Maximum(int startRow, int startColumn, int endRow, int endColumn)
    {
        int max = -2147483647;
        int maxValue = max;
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            max = maxValue;
        }

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = getCurrentData(currentCol->data);
                if (data > max)
                {
                    max = data;
                }
            }
        }

        if (max != maxValue)
        {
            current->data = to_string(max);
        }
    }

    vector<T> cellsData;
    void Copy(int startRow, int startColumn, int endRow, int endColumn)
    {
        char pie = 227;
        string str(1, pie);
        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                cellsData.push_back(currentCol->data);
            }
            cellsData.push_back(str);
        }
    }

    void Cut(int startRow, int startColumn, int endRow, int endColumn)
    {
        char pie = 227;
        string str(1, pie);
        T sum = T();
        for (int row = startRow; row <= endRow; row++)
        {
            Node<T>* currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T>* currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                cellsData.push_back(currentCol->data);
                currentCol->data = "     ";
            }
            cellsData.push_back(str);
        }
    }

    void Paste()
    {
        char pie = 227;
        string str(1, pie);
        Node<T>* node = current;
        Node<T>* temp = current;
        for (int i = 0; i < cellsData.size() - 1; i++)
        {
            if (cellsData[i] != str)
            {
                if (node->right == nullptr && cellsData[i + 1] != str)
                {
                    InsertColumnAtRight();
                }
                node->data = cellsData[i];
                node = node->right;
            }
            else
            {
                InserRowBelow();
                node = temp->down;
                temp = node;
            }
        }
    }

    Coordinate GetRowandColumn()
    {
        int row = -1, col = -1;
        Node<T>* temp = root;
        Node<T>* first = root;
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                if (temp == current)
                {
                    row = i;
                    col = j;
                    break;
                }
                temp = temp->right;
            }
            if (first->down != nullptr)
            {
                first = first->down;
            }
            temp = first;
        }
        Coordinate C(row, col);
        return C;
    }
};
/*
1.	VK_LEFT: To move the current to the left.
2.	VK_RIGHT : To move the current to the right.
3.	VK_UP : To move the current up.
4.	VK_DOWN : To move the current down.
5.	VK_SHIFT + W : To insert a row above the current.
6.	VK_SHIFT + S : To insert a row below the current.
7.	VK_SHIFT + A : To insert a column left to the current.
8.	VK_SHIFT + D : To insert a column right to the current.
9.	VK_SHIFT + VK_OEM_COMMA : To select the lower range for operation.
10.	VK_SHIFT + VK_OEM_PERIOD : To select the upper range for operation.
11.	VK_SHIFT + C : To copy the content from the given range.
12.	VK_SHIFT + X : To cut the content from the given range.
13.	VK_SHIFT + V : To paste the copied or cut content starting from the current.
14.	VK_SHIFT + I : To give input in the current.
15.	VK_SHIFT + Q : To delete the current column.
16.	VK_SHIFT + R : To delete the current row.
17.	Q : To clear all the content of the current column.
18.	R : To clear all the content of the current row.
19.	VK_F1 : To insert the cell by right shift.
20.	VK_F2 : To insert the cell by downshift.
21.	VK_F3 : To delete the cell by left shift.
22.	VK_F4 : To delete the cell by upshift.
23.	VK_F5 : To get the range sum in a given range and set it in the current node.
24.	VK_F6 : To get the range average in a given range and set it in the current node.
25.	VK_F7 : To count the number of nodes in a given range and set it in the current node.
26.	VK_F8 : To get the minimum value in a given range and set it in the current node.
27.	VK_F9 : To get maximum value in a given range and set it in the current node.
*/

int main()
{
    int startRow, startCol;
    int endRow, endCol;
    MiniExcel<string> ME(5, 5);
    string option = "";
    system("cls");
    ClearScreen();
    Coordinate C1(-1, -1);
    Coordinate C2(-1, -1);
    while (true)
    {
        hideCursor();
        ClearScreen();
        ME.PrintGrid();
        if (GetAsyncKeyState(VK_LEFT) && ME.GetCurrentCell()->left != nullptr)
        {
            ME.SetCurrentToLeft();
        }

        if (GetAsyncKeyState(VK_RIGHT) && ME.GetCurrentCell()->right != nullptr)
        {
            ME.SetCurrentToRight();
        }

        if (GetAsyncKeyState(VK_UP) && ME.GetCurrentCell()->up != nullptr)
        {
            ME.SetCurrentToUp();
        }

        if (GetAsyncKeyState(VK_DOWN) && ME.GetCurrentCell()->down != nullptr)
        {
            ME.SetCurrentToDown();
        }

        if ((GetAsyncKeyState('W') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            ME.InserRowAbove();
        }

        if ((GetAsyncKeyState('S') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            ME.InserRowBelow();
        }

        if ((GetAsyncKeyState('A') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            ME.InsertColumnToLeft();
        }

        if ((GetAsyncKeyState('D') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            ME.InsertColumnToRight();
        }

        if ((GetAsyncKeyState(VK_OEM_COMMA) & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            C1 = ME.GetRowandColumn();
        }

        if ((GetAsyncKeyState(VK_OEM_PERIOD) & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            C2 = ME.GetRowandColumn();
        }

        if ((GetAsyncKeyState('C') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Copy(startRow, startCol, endRow, endCol);
            }
        }

        if ((GetAsyncKeyState('X') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Cut(startRow, startCol, endRow, endCol);
            }
        }

        if ((GetAsyncKeyState('V') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Paste();
            }
        }

        if ((GetAsyncKeyState('I') & 0x8000) && (GetAsyncKeyState(VK_SHIFT) & 0x8000))
        {
            ME.TakeInput();
        }

        if ((GetAsyncKeyState('Q') & 0x8000) && (GetAsyncKeyState(VK_SHIFT)))
        {
            ME.DeleteColumn();
            system("cls");
        }

        if (GetAsyncKeyState('R') & 0x8000 && (GetAsyncKeyState(VK_SHIFT)))
        {
            ME.DeleteRow();
            system("cls");
        }

        if ((GetAsyncKeyState('R') & 0x8000))
        {
            ME.ClearRow();
        }

        if ((GetAsyncKeyState('Q') & 0x8000))
        {
            ME.ClearColumn();
        }

        if ((GetAsyncKeyState(VK_F1) & 0x8000))
        {
            ME.InsertCellByRightShift();
        }

        if ((GetAsyncKeyState(VK_F2) & 0x8000))
        {
            ME.InsertCellByDownShift();
        }

        if ((GetAsyncKeyState(VK_F3) & 0x8000))
        {
            ME.DeleteCellByLeftShift();
        }

        if ((GetAsyncKeyState(VK_F4) & 0x8000))
        {
            ME.DeleteCellByUpShift();
        }

        if ((GetAsyncKeyState(VK_F5) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                int sum = ME.GetRangeSum(startRow, startCol, endRow, endCol);
                ME.SetCurrentData(to_string(sum));
            }
        }

        if ((GetAsyncKeyState(VK_F6) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                int avg = ME.GetRangeAverge(startRow, startCol, endRow, endCol);
                ME.SetCurrentData(to_string(avg));
            }
        }

        if ((GetAsyncKeyState(VK_F7) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Count(startRow, startCol, endRow, endCol);
            }
        }

        if ((GetAsyncKeyState(VK_F8) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Minimum(startRow, startCol, endRow, endCol);
            }
        }

        if ((GetAsyncKeyState(VK_F9) & 0x8000))
        {
            startRow = C1.row;
            startCol = C1.col;
            endRow = C2.row;
            endCol = C2.col;
            if (startRow != -1 && startCol != -1 && endRow != -1 && endCol != -1)
            {
                ME.Maximum(startRow, startCol, endRow, endCol);
            }
        }
        Sleep(100);
    }
    return 0;
}

void ClearScreen()
{
    COORD cursorPosition;
    cursorPosition.X = 0;
    cursorPosition.Y = 0;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), cursorPosition);
}

void hideCursor()
{
    CONSOLE_CURSOR_INFO cursorInfo;
    cursorInfo.dwSize = 100;
    cursorInfo.bVisible = FALSE;
    SetConsoleCursorInfo(GetStdHandle(STD_OUTPUT_HANDLE), &cursorInfo);
}