#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->lineEdit->setText(QStringLiteral("C:/Users/MECHREVO/Desktop/材料出库明细表.xlsx"));
}

MainWindow::~MainWindow()
{
    delete ui;
}


//点击选择目录
void MainWindow::on_toolButton_clicked()
{
    QString directory = QFileDialog::getOpenFileName(this,QStringLiteral("选择要读取的Excel文件"),"",tr("Excel(*.xlsx)"));        //选择路径
    if (!directory.isEmpty())
    {
        ui->lineEdit->setText(directory);
    }
}
//解析Excel
void MainWindow::on_pushButton_jiexi_clicked()
{
    QXlsx::Document doc(ui->lineEdit->text());                                      //读取文件
    int rowCounts = doc.dimension().lastRow ();                                     //获取打开文件的最后一行（注意，如果最后一行有空格也为有效行）
    int colCounts = doc.dimension().lastColumn ();                                  //获取打开文件的最后一列
    if (!doc.load())
        {
        QMessageBox::information(
                    this,
                    QStringLiteral("错误"),
                    QStringLiteral("文件打开失败！请检查文件或路径！"),
                    QStringLiteral("确定"));
        return;
    }

    //读取部门信息
    for(int i=3; i <=rowCounts;++i )
    {
        qDebug() <<  QStringLiteral("读取部门第 %1 遍").arg(i);
        if(doc.read( i, 1 ).toString() != ""&&list_bumen.contains(doc.read( i, 1 ).toString()) == false)//获取表格不为空
        {
            qDebug() <<  QStringLiteral("加入部门数据 : %1").arg(doc.read( i, 1 ).toString());
            list_bumen.append(doc.read( i, 1 ).toString());
        }else{
            qDebug() <<  QStringLiteral("跳过部门数据 : %1").arg(doc.read( i, 1 ).toString());
        }
    }
    //读取物品信息
    for(int i=3; i <=rowCounts;++i )
    {
        qDebug() <<  QStringLiteral("读取物品第 %1 遍").arg(i);
        if(doc.read( i, 3 ).toString() != ""&&list_wupin.contains(doc.read( i, 3 ).toString()) == false)//获取表格不为空
        {
            qDebug() <<  QStringLiteral("加入物品数据 : %1").arg(doc.read( i, 3 ).toString());
            list_wupin.append(doc.read( i, 3 ).toString());         //加入进物品列表

            //  加入单价信息
            if(!Wupin_Map.contains(doc.read( i, 3 ).toString()))
            {
                Wupin_Map.insert(doc.read( i, 3 ).toString(), doc.read( i, 6 ).toDouble()); //向map里添加一对“键-值”
                qDebug() <<QStringLiteral("加入 %1 单价 : %2").arg(doc.read( i, 3 ).toString()).arg(doc.read( i, 6 ).toDouble());
            }
            else {
                qDebug() <<  QStringLiteral("%1物品单价重复").arg(doc.read( i, 3 ).toString());
            }

        }else{
            qDebug() <<  QStringLiteral("跳过物品数据 : %1").arg(doc.read( i, 3 ).toString());
        }
    }
    qDebug() << QStringLiteral("部门数据共 %1 条：%2").arg(list_bumen.size()).arg(list_bumen.join(","));
    qDebug() << QStringLiteral("物品数据共 %1 条：%2").arg( list_wupin.size()).arg(list_wupin.join(","));
    qDebug() << QStringLiteral("最大行数：%1 最大列数：%2").arg(rowCounts).arg(colCounts);

    QXlsx::Document xlsxDoc;                                                                            //创建新文档
    //QXlsx::Format title_format;                                                                       //设置格式
    //title_format.setFontSize(11);                                                                     //字体大小
    //title_format.setFontBold(true);                                                               //加粗
    //title_format.setFontColor(QColor(Qt::red));                                           //设置字体颜色 red(红色) white(白色) darkBlue(深蓝) QColor("#EACC93")
    //title_format.setBorderStyle(QXlsx::Format::BorderThin);                       //边框
    //title_format.setHorizontalAlignment(QXlsx::Format::AlignLeft);           //AlignLeft(左对齐),AlignHCenter(中心对齐)
    //title_format.setVerticalAlignment(QXlsx::Format::AlignVCenter);        //垂直居中

    //title_format.setFillPattern(QXlsx::Format::PatternSolid);
    //title_format.setPatternBackgroundColor(Qt::darkBlue);     //设置背景色

    //xlsxDoc.setRowHeight(1, 80);                                   //设置行高 第一行
    //xlsxDoc.setColumnWidth(day*2, day*2, 13);             //设置列宽

    //循环写入部门
    for (int i = 3;i <=  list_bumen.size()+2; ++i) {
        QXlsx::Format BumenStyle;       //创建部门样式模板
        BumenStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);     //中心对齐
        xlsxDoc.setColumnWidth(1, 1, 13);           //设置列宽
        qDebug() << QStringLiteral("部门%1 ID：%2").arg(list_bumen.at(i-3)).arg(i-3);
        xlsxDoc.write(i, 1, list_bumen.at(i-3),BumenStyle);        //写入数据
    }
    //循环写入物品
    for (int i = 2, j = 0;j <  list_wupin.size(); ++j) {
        QXlsx::Format WupinStyle;   //创建物品样式模板
        WupinStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);     //中心对齐

        xlsxDoc.write(1, i, list_wupin.at(j));
        qDebug() << QStringLiteral("ID：%1 物品名：%2 单价：%3").arg(j).arg(list_wupin.at(j)).arg(Wupin_Map[list_wupin.at(j)]);
        xlsxDoc.mergeCells(QXlsx::CellRange(1, i, 1, i+1),WupinStyle);       //合并单元格
        i = i+2;
    }
    xlsxDoc.saveAs("datetime.xlsx");
}

