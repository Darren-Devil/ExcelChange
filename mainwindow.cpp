#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QMessageBox>
#include "qexcel.h"
#include <QApplication>
#include <QDebug>
#include <QElapsedTimer>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->openBtn, SIGNAL(clicked()), this, SLOT(onOpenBtnClicked()));
    connect(ui->okBtn, SIGNAL(clicked()), this, SLOT(onOkBtnClicked()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

/*********************************************************************
*     function:openExcel()
*  Description:打开Excel文件
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
void MainWindow::openExcel()
{ 
    QExcel fileName(QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), "*.xlsx"));
    QElapsedTimer timer;
    timer.start();
    qDebug()<<fileName.getSheetsCount();
    fileName.selectSheet(5);
    qDebug()<<"open cost:"<<timer.elapsed()<<"ms";
}

void MainWindow::onOpenBtnClicked()
{
    openExcel();
}
/*********************************************************************
*     function:onOkBtnClicked()
*  Description:按下确认键修改Excel文件
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
void MainWindow::onOkBtnClicked()
{
    QExcel file("*.xlsx");
    QElapsedTimer timer;
    timer.start();
    qDebug()<<file.getSheetsCount();
    qDebug()<<"open cost:"<<timer.elapsed()<<"ms";timer.restart();
    file.selectSheet(5);
    file.clearCell(50,2);
    file.setCellString(50,2,ui->changeEdit->text());
    file.save();
    qDebug()<<"change data cost:"<<timer.elapsed()<<"ms";
//    file.close();
}
