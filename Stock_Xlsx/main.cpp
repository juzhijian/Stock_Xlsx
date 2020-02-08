#include "form/mainwindow.h"
#include "form/quiwidget.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    a.setFont(QFont("Microsoft Yahei", 9));//设置字体
    a.setWindowIcon(QIcon(":/main.ico"));//设置软件图标

    QUIWidget w;
    MainWindow *creator = new MainWindow;

    //设置主窗体
    w.setMainWidget(creator);
    //QObject::connect(&w, SIGNAL(changeStyle(QString)), creator, SLOT(changeStyle(QString)));

    //设置标题
    w.setTitle("Excel 解析");

    //设置标题文本居中
    //w.setAlignment(Qt::AlignCenter);

    //设置窗体可拖动大小
    w.setSizeGripEnabled(true);

    //设置换肤下拉菜单可见
    w.setVisible(QUIWidget::BtnMenu, false);

    //设置标题栏高度
    //w.setTitleHeight(50);

    //设置按钮宽度
    //w.setBtnWidth(50);

    //设置左上角图标-图形字体
    //w.setIconMain(QChar(0xf099), 11);

    //设置左上角图标-图片文件
    //w.setPixmap(QUIWidget::Lab_Ico, ":/main.ico");

    //在main函数中加载qss文件
    QFile file(":/qss/flatwhite.css");
    if (file.open(QFile::ReadOnly))
    {
        QString stylesheet = QLatin1String(file.readAll());
        qApp->setStyleSheet(stylesheet);
        file.close();
    }
    else
    {
        QMessageBox::warning(NULL, "warning", "openfailed",QMessageBox::Yes | QMessageBox::No,QMessageBox::Yes);
    }

    w.show();
    return a.exec();
}
