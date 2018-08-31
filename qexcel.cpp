#include <QAxObject>
#include <QFile>
#include <QStringList>
#include <QDebug>
#include <QFileDialog>

#include "qexcel.h"

/*********************************************************************
*     function:QExcel(QString xlsxFilePath, QObject *parent)
*  Description:构造函数打开Excel文件
*        Input:
*       Output:
*       Return:
*       Others:
*       Auther:dongyuchuan
*  Create Time:2018.08.28
*---------------------------------------------------------------------
*  Modify
*  Version       Author        Date           Modification
*  V0.00         dongyuchuan   2018.08.28     实现函数功能
*
**********************************************************************/
QExcel::QExcel(QString xlsxFilePath, QObject *parent)
{
    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;

    excel = new QAxObject("Excel.Application", parent);
    excel->dynamicCall("SetVisible(bool)", true);
    workBooks = excel->querySubObject("Workbooks");
    /*
//    QFile file(xlsxFilePath);
//    if (file.exists())
//    {
*/
        workBooks->dynamicCall("Open(const QString&)", xlsxFilePath);
        workBook = excel->querySubObject("ActiveWorkBook");
        sheets = workBook->querySubObject("WorkSheets");
        /*
//    }
*/
}

QExcel::~QExcel()
{
    close();
}
/*********************************************************************
*     function:close()
*  Description:关闭Excel文件
*        Input:
*       Output:
*       Return:
*       Others:
*       Auther:dongyuchuan
*  Create Time:2018.08.28
*---------------------------------------------------------------------
*  Modify
*  Version       Author        Date           Modification
*  V0.00         dongyuchuan   2018.08.28     实现函数功能
*
**********************************************************************/
void QExcel::close()
{
    excel->dynamicCall("Quit()");

    delete sheet;
    delete sheets;
    delete workBook;
    delete workBooks;
    delete excel;

    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;
}

QAxObject *QExcel::getWorkBooks()
{
    return workBooks;
}

QAxObject *QExcel::getWorkBook()
{
    return workBook;
}

QAxObject *QExcel::getWorkSheets()
{
    return sheets;
}

QAxObject *QExcel::getWorkSheet()
{
    return sheet;
}

void QExcel::selectSheet(const QString& sheetName)
{
    sheet = sheets->querySubObject("Item(const QString&)", sheetName);
}

void QExcel::deleteSheet(const QString& sheetName)
{
    QAxObject * a = sheets->querySubObject("Item(const QString&)", sheetName);
    a->dynamicCall("delete");
}

void QExcel::deleteSheet(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    a->dynamicCall("delete");
}

void QExcel::selectSheet(int sheetIndex)
{
    sheet = sheets->querySubObject("Item(int)", sheetIndex);
}

void QExcel::setCellString(int row, int column, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QExcel::mergeCells(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

void QExcel::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    QString cell;
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(":");
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

QVariant QExcel::getCellValue(int row, int column)
{
    QVariant data;
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    //        return range->property("value");
    if ( range )
    {
        data = range->dynamicCall("Value2()");
    }
    return data;
}

void QExcel::save()
{
    workBook->dynamicCall("Save()");
}

int QExcel::getSheetsCount()
{
    return sheets->property("Count").toInt();
}

QString QExcel::getSheetName()
{
    return sheet->property("Name").toString();
}

QString QExcel::getSheetName(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    return a->property("Name").toString();
}

void QExcel::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject *rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject *columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}

void QExcel::insertSheet(QString sheetName)
{
    sheets->querySubObject("Add()");
    QAxObject * a = sheets->querySubObject("Item(int)", 1);
    a->setProperty("Name", sheetName);
}

void QExcel::clearCell(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QExcel::clearCell(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

int QExcel::getUsedRowsCount()
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    int topRow = usedRange->property("Row").toInt();
    QAxObject *rows = usedRange->querySubObject("Rows");
    int bottomRow = topRow + rows->property("Count").toInt() - 1;
    return bottomRow;
}

void QExcel::setCellString(const QString& cell, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("SetValue(const QString&)", value);
}
