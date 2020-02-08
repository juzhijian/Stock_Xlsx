#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "quiwidget.h"

MainWindow::MainWindow(QWidget *parent)
	: QMainWindow(parent)
	, ui(new Ui::MainWindow)
{
	ui->setupUi(this);
	this->initForm();
	/*
	ui->lineEdit->setText(QStringLiteral("C:/Users/MECHREVO/Desktop/材料出库明细表.xlsx"));		                               //测试写入
	*/
}

MainWindow::~MainWindow()
{
	delete ui;
}

void MainWindow::initForm()
{
	//引入图形字体
	int fontId = QFontDatabase::addApplicationFont(":/image/fontawesome-webfont.ttf");
	QString fontName = QFontDatabase::applicationFontFamilies(fontId).at(0);
	iconFont = QFont(fontName);

	//QTimer::singleShot(100, this, SLOT(initPanelWidget()));
}



/*点击选择目录*/
void MainWindow::on_toolButton_clicked()
{
	/*选择路径*/
	QString directory = QFileDialog::getOpenFileName(this, "选择要读取的Excel文件", "", tr("Excel(*.xlsx)"));
	if (!directory.isEmpty())
	{
		ui->lineEdit->setText(directory);
	}
}

/*解析Excel*/
void MainWindow::on_pushButton_jiexi_clicked()
{
	/*打开文件*/
	QXlsx::Document doc(ui->lineEdit->text());                                                                             //读取文件
	int rowCounts = doc.dimension().lastRow();                                                                             //获取打开文件的最后一行（注意，如果最后一行有空格也为有效行）
	int colCounts = doc.dimension().lastColumn();                                                                          //获取打开文件的最后一列

	if (!doc.load())
	{
		QMessageBox::information(this, "错误", "文件打开失败！请检查文件或路径！", "确定");
		return;
	}

	qDebug() << QString("最大行数：%1 最大列数：%2").arg(rowCounts).arg(colCounts);


	/*读取操作*/

	/*读取部门信息*/
	for (int i = 3; i <= rowCounts; ++i)
	{
		qDebug() << QString("读取部门第 %1 遍").arg(i);
		if (doc.read(i, 1).toString() != ""&&list_bumen.contains(doc.read(i, 1).toString()) == false)                      //获取表格不为空
		{
			qDebug() << QString("加入部门数据 : %1").arg(doc.read(i, 1).toString());
			list_bumen.append(doc.read(i, 1).toString());
		}
		else
		{
			qDebug() << QString("跳过部门数据 : %1").arg(doc.read(i, 1).toString());
		}
	}

	/*创建json用于储存物品信息 名称 单位 价格*/
	QJsonObject Wupin_Object;
	QJsonObject Wupin_data_Object;

	/*读取物品信息*/
	for (int i = 3, k = 0; i <= rowCounts; ++i)
	{
		qDebug() << QString("读取物品第 %1 遍").arg(k);

		if (doc.read(i, 3).toString() != ""
			&&list_wupin.contains(doc.read(i, 3).toString()) == false)                                                     //获取表格不为空 且 物品列表无
		{
			/*向物品列表写入数据*/
			list_wupin.append(doc.read(i, 3).toString());                                                                  //加入进物品列表
			qDebug() << QString("加入物品数据 : %1").arg(doc.read(i, 3).toString());

			/*写入map 记录 物品名称 单价*/
			Wupin_Map.insert(doc.read(i, 3).toString(), doc.read(i, 6).toDouble());                                        //向map里添加一对“物品吗-单价”
			qDebug() << QString("加入 %1 单价 : %2").arg(doc.read(i, 3).toString()).arg(doc.read(i, 6).toDouble());

			/*写入json*/
			Wupin_data_Object.insert("name", doc.read(i, 3).toString());                                                   //物品名称
			Wupin_data_Object.insert("unit", doc.read(i, 4).toString());                                                   //物品单位
			Wupin_data_Object.insert("UnitPrice", doc.read(i, 6).toDouble());                                              //单价
			Wupin_Object.insert(QString::number(k), Wupin_data_Object);

			++k;
		}
		else if (list_wupin.contains(doc.read(i, 3).toString()) == true
			&& Wupin_Map[doc.read(i, 3).toString()] != doc.read(i, 6))                                                     //搜索物品列表是否存在 且 价格不同
		{
			ui->textBrowser->append(QString("第 %1 行，%2物品单价重复！").arg(i).arg(doc.read(i, 3).toString()));             //文本框提示消息
			qDebug() << Wupin_Map[doc.read(i, 3).toString()] << "x" << doc.read(i, 6).toString();
		}
		else if (doc.read(i, 3).toString() == "")                                                                          //表格数据空
		{
			qDebug() << QString("空物品数据");
		}
		else
		{
			qDebug() << QString("重复物品数据 : %1").arg(doc.read(i, 3).toString());
		}
	}

	/*创建json用于储存单据*/
	QJsonObject Danju_Object;
	QJsonObject Danju_data_Object;

	/*读取单据数据*/
	for (int i = 3; i < rowCounts; ++i)
	{
		if (doc.read(i, 1).toString() != ""
			&&Wupin_Map[doc.read(i, 3).toString()] == doc.read(i, 6))                                                      //单据行不为空 且 单价相同
		{
			/*从josn获取当前单据数据*/
			Danju_data_Object = Danju_Object.value(doc.read(i, 1).toString()).toObject();                                  //当前单据数据

			/*单据当前物品信息不存在，则写入当前读取信息*/
			if (Danju_data_Object[doc.read(i, 3).toString()].isNull())
			{
				Danju_data_Object.insert(doc.read(i, 3).toString(), doc.read(i, 5).toDouble());                            //写入单据数据 物品名称 数量
				Danju_Object.insert(doc.read(i, 1).toString(), Danju_data_Object);
			}
			/*单据当前物品信息不为空，则原始数量+当前读取数量*/
			else if (!Danju_data_Object[doc.read(i, 3).toString()].isNull())
			{
				Danju_data_Object.insert(doc.read(i, 3).toString(),
					doc.read(i, 5).toDouble() + Danju_data_Object[doc.read(i, 3).toString()].toDouble());                  //写入单据数据 物品名称 原始数量+读取数量
				Danju_Object.insert(doc.read(i, 1).toString(), Danju_data_Object);
			}
		}
		else if (doc.read(i, 1).toString() != ""
			&&Wupin_Map[doc.read(i, 3).toString()] != doc.read(i, 6))
		{
			qDebug() << QString("%1 单价不同 记录值：%2 读取值：%3").arg(doc.read(i, 3).toString()).arg(Wupin_Map[doc.read(i, 3).toString()]).arg(doc.read(i, 6).toString());
			ui->textBrowser->append(QString("第 %1 行，%2 单价不同 记录值：%3 读取值：%4").arg(i).arg(doc.read(i, 3).toString()).arg(Wupin_Map[doc.read(i, 3).toString()]).arg(doc.read(i, 6).toString()));
		}
		else if (doc.read(i, 1).toString() == "")
		{
			qDebug() << QString("空单据数据");
		}
	}

	qDebug() << Wupin_Object;
	qDebug() << Danju_Object;
	qDebug() << QString("部门数据共 %1 条：%2").arg(list_bumen.size()).arg(list_bumen.join(","));
	qDebug() << QString("物品数据共 %1 条：%2").arg(list_wupin.size()).arg(list_wupin.join(","));


	/*设置样式*/

	QXlsx::Document xlsxDoc;                                                                                               //创建新文档
	xlsxDoc.setColumnWidth(1, 13);                                                                                         //设置列宽

	QXlsx::Format BumenStyle;                                                                                              //创建部门样式模板
	BumenStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);                                                        //中心对齐

	QXlsx::Format WupinStyle;                                                                                              //创建物品样式模板
	WupinStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);                                                        //中心对齐

	QXlsx::Format NeirongStyle;                                                                                            //创建内容模板
	NeirongStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);                                                      //中心对齐

	/*
	QXlsx::Format title_format;                                                                                            //设置格式
	title_format.setFontSize(11);                                                                                          //字体大小
	title_format.setFontBold(true);                                                                                        //加粗
	title_format.setFontColor(QColor(Qt::red));                                                                            //设置字体颜色 red(红色) white(白色) darkBlue(深蓝) QColor("#EACC93")
	title_format.setBorderStyle(QXlsx::Format::BorderThin);                                                                //边框
	title_format.setHorizontalAlignment(QXlsx::Format::AlignLeft);                                                         //AlignLeft(左对齐),AlignHCenter(中心对齐)
	title_format.setVerticalAlignment(QXlsx::Format::AlignVCenter);                                                        //垂直居中

	title_format.setFillPattern(QXlsx::Format::PatternSolid);
	title_format.setPatternBackgroundColor(Qt::darkBlue);                                                                  //设置背景色

	xlsxDoc.setRowHeight(1, 80);                                                                                           //设置行高 第一行
	xlsxDoc.setColumnWidth(day*2, day*2, 13);                                                                              //设置列宽
	*/

	/*单独写入*/
	xlsxDoc.write(1, 1, QString("品    名"), NeirongStyle);
	xlsxDoc.write(2, 1, QString("单位单价"), NeirongStyle);
	xlsxDoc.write(3, 1, QString("领用部门"), NeirongStyle);

	/*循环写入部门*/
	for (int i = 0; i < list_bumen.size(); ++i) {
		xlsxDoc.write(i + 4, 1, list_bumen.at(i), BumenStyle);                                                             //写入数据
		qDebug() << QString("部门%1 ID：%2").arg(list_bumen.at(i)).arg(i);
	}

	int row = xlsxDoc.dimension().lastRow();                                                                               //获取打开文件的最后一行（注意，如果最后一行有空格也为有效行）

	/*循环写入物品*/
	for (int i = 2, j = 0; j < list_wupin.size(); ++j) {

		/*临时读取当前物品信息*/
		QJsonObject json_wupin;
		json_wupin = Wupin_Object.value(QString::number(j)).toObject();                                                    //当前物品信息

		/*写入数据*/
		xlsxDoc.write(1, i, json_wupin["name"].toString());                                                                //写入物品 行，列，内容，样式
		xlsxDoc.write(2, i, json_wupin["unit"].toString(), NeirongStyle);                                                  //写入价格 行，列，内容，样式
		xlsxDoc.write(2, i + 1, json_wupin["UnitPrice"].toDouble(), NeirongStyle);                                         //写入价格 行，列，内容，样式
		xlsxDoc.write(3, i, QString("数量"), NeirongStyle);                                                           //写入数量 行，列，内容，样式
		xlsxDoc.write(3, i + 1, QString("金额"), NeirongStyle);                                                       //写入金额 行，列，内容，样式

		/*写入合计 行，列，内容，样式*/
		xlsxDoc.write(row + 1, i, QString("=SUM(%1)").arg(QXlsx::CellRange(4, i, row, i).toString(false)), NeirongStyle);
		/*写入合计 行，列，内容，样式*/
		xlsxDoc.write(row + 1, i + 1, QString("=SUM(%1)").arg(QXlsx::CellRange(4, i + 1, row, i + 1).toString(false)), NeirongStyle);

		/*循环设置单元格样式*/
		xlsxDoc.mergeCells(QXlsx::CellRange(1, i, 1, i + 1), WupinStyle);                                                  //合并单元格
		xlsxDoc.setColumnWidth(i, 6);                                                                                      //设置列宽
		xlsxDoc.setColumnWidth(i + 1, 9);                                                                                  //设置列宽
		i = i + 2;

		qDebug() << QString("ID：%1 物品名：%2 单价：%3").arg(j).arg(list_wupin.at(j)).arg(Wupin_Map[list_wupin.at(j)]);
	}

	/*循环写入物品数量*/
	for (int i = 4, j = 0; j < list_bumen.size(); ++j)
	{
		/*临时读取当前物品信息*/
		QJsonObject json_Danju;
		json_Danju = Danju_Object.value(list_bumen.at(j)).toObject();                                                      //当前物品信息
		qDebug() << list_bumen.at(j) << json_Danju;
		int Bumen_row = i;
		for (int i = 2, k = 0; k < list_wupin.size(); ++k)
		{
			if (!json_Danju[xlsxDoc.read(1, i).toString()].isNull())
			{
				//qDebug() << "写入" << xlsxDoc.read(1, i).toString() << "数量";
				/*写入数量*/
				if (ui->checkBox->isChecked() == true)
				{
					xlsxDoc.write(Bumen_row, i, json_Danju[xlsxDoc.read(1, i).toString()].toDouble(), NeirongStyle);     //写入数量 行，列，内容，样式
				}

				/*写入价格*/
				if (ui->checkBox_2->isChecked() == true)
				{
					xlsxDoc.write(Bumen_row, i + 1, json_Danju[xlsxDoc.read(1, i).toString()].toDouble()*xlsxDoc.read(2, i + 1).toDouble(), NeirongStyle);
				}
			}
			i = i + 2;
		}
		i = i + 1;
	}

	qDebug() << QString("解析成功");
	ui->textBrowser->append(QString("解析成功"));
	/*保存文档*/
	xlsxDoc.saveAs("datetime.xlsx");
}
