#ifndef QEXCEL_H
#define QEXCEL_H

#include <QString>
#include <QVariant>

class QAxObject;

class QExcel : public QObject
{
public:
    QExcel(QString xlsFilePath, QObject *parent = 0);//打开工作表集合
    ~QExcel();

public:
    QAxObject * getWorkBooks();
    QAxObject * getWorkBook();
    QAxObject * getWorkSheets();
    QAxObject * getWorkSheet();

public:
    void selectSheet(const QString& sheetName);
    //sheetIndex 起始于 1
    //选择工作表
    void selectSheet(int sheetIndex);
    void deleteSheet(const QString& sheetName);
    void deleteSheet(int sheetIndex);
    void insertSheet(QString sheetName);
    int getSheetsCount();
    //在 selectSheet() 之后才可调用
    QString getSheetName();
    QString getSheetName(int sheetIndex);

    void setCellString(int row, int column, const QString& value);
    //cell
    void setCellString(const QString& cell, const QString& value);
    //range
    void mergeCells(const QString& range);
    void mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn);
    QVariant getCellValue(int row, int column);
    void clearCell(int row, int column);
    void clearCell(const QString& cell);

    void getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn);
    int  getUsedRowsCount();

    void save();
    void close();

private:
    QAxObject * excel;
    QAxObject * workBooks;
    QAxObject * workBook;
    QAxObject * sheets;
    QAxObject * sheet;
};

#endif
