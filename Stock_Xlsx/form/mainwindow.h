#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QFileDialog>          //打开文件
#include <QDebug>               //调试消息
#include <QMessageBox>          //显示提示框
#include <QtCore>

//Qxlsx类
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_toolButton_clicked();

    void on_pushButton_jiexi_clicked();

private slots:
    void initForm();

private:
    Ui::MainWindow *ui;
    QFont iconFont;             //图形字体
    QMap<QString,double> Wupin_Map;                //定义一个物品列表 物品名称 价格
    QStringList list_bumen;                        //定义一个部门的列表
    QStringList list_wupin;                        //定义一个物品的列表
};
#endif // MAINWINDOW_H
