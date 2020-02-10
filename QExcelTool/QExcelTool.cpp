#include "QExcelTool.h"
#include <QFile>
#include <QMessageBox>
#include <QFileDialog>

int err = 0, ftype = -1;
QStatusBar* statbar = nullptr;
std::string content;
QFile* xfile = nullptr;

QExcelTool::QExcelTool(QWidget* parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	ui.statusBar->setSizeGripEnabled(false);
	fxls = nullptr;
	connect(ui.pushButton_1, SIGNAL(clicked()), this, SLOT(openxlsfile()));
	statbar = ui.statusBar;
	ui.statusBar->showMessage(QString::fromLocal8Bit("就绪"));
}

void QExcelTool::openxlsfile()
{
	fpath = QFileDialog::getOpenFileName(this, QString::fromLocal8Bit("打开表格文件"), ".", QString::fromLocal8Bit("表格文件(*.csv)")).replace("/", "\\");
	ui.textBrowser->setText(fpath);
	if (fpath.isEmpty())
	{
		ui.statusBar->showMessage(QString::fromLocal8Bit("错误：未选择文件"));
		err = QETERROR_NOFILESELECTED;
	}
	else
	{
		fxls = new QFile(fpath);
		fxls->open(QIODevice::ReadOnly | QIODevice::Text);
		if (!fxls->isOpen())
		{
			ui.statusBar->showMessage(QString::fromLocal8Bit("错误：打开文件失败（%1)").arg(fxls->errorString()));
			err = QETERROR_FILEOPENFAIL;
		}
		else
		{
			std::string spath;
			spath = fpath.toStdString();
			spath = spath.substr(spath.find_last_of('\\') + 1);
			statbar->showMessage(QString::fromLocal8Bit("信息：打开文件成功\t文件名：\"%1\"").arg(QString(spath.c_str())));
			spath = fpath.toStdString();
			spath = spath.substr(spath.find_last_of('.') + 1);
			ftype = (spath == "csv" ? QETFILETYPE_CSV : (spath == "xls" ? QETFILETYPE_XLS : (spath == "xlsx" ? QETFILETYPE_XLSX : -1)));
			if (ftype == -1)
			{
				QMessageBox::critical(this, "ERROR", "Unknown Error");
				QApplication::quit();
			}
			err = QETERROR_NOERROR;
			xfile = fxls;
			CreateThread(NULL, 0, ReadFile, NULL, NULL, NULL);
		}
	}
}

DWORD WINAPI ReadFile(LPVOID param)
{
	statbar->showMessage(QString::fromLocal8Bit("信息：读取文件..."));
	content = QString(xfile->readAll()).toStdString();
	qDebug(content.c_str());
	statbar->showMessage(QString::fromLocal8Bit("就绪"));
	return 0;
}
