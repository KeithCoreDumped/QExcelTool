#include "QExcelTool.h"
#include <QFile>
#include <QMessageBox>
#include <QStandardItemModel>
#include <QFileDialog>
#include <QTextStream>
#include <QTextCodec>

int err = 0, ftype = -1, ntvline = 0, ntvrow = 0, xdec = 0;
QStatusBar* xstatusBar = nullptr;
QFile* xfile = nullptr;
QTableView* xtableView = nullptr;
QString xpath, xcontent;
QPushButton* xqpb;
QStandardItemModel* model;

/******************************************************************************************************/
QExcelTool::QExcelTool(QWidget* parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	ui.statusBar->setSizeGripEnabled(false);
	ui.tableView->setShowGrid(true);
	xfile = nullptr;
	xqpb = ui.ChangeDecButton;
	xtableView = ui.tableView;
	connect(ui.pushButton_1, SIGNAL(clicked()), this, SLOT(OpenFile()));
	connect(ui.ChangeDecButton, SIGNAL(clicked()), this, SLOT(ChangeDec()));
	xstatusBar = ui.statusBar;
	ui.statusBar->showMessage(QString::fromLocal8Bit("就绪"));
	ui.textBrowser->setToolTipDuration(1000);
	model = new QStandardItemModel();
	//QStringList labels = QString::fromLocal8Bit("1,2,3").simplified().split(",");
	//model->setHorizontalHeaderLabels(labels);
	//ui.tableView->setModel(model);
	//ui.tableView->show();
}
/******************************************************************************************************/
inline bool isempty(QString src)
{
	return src.isNull() || src.isEmpty();
}
/******************************************************************************************************/
QString getlinename(unsigned int i)
{
	///STACK OVERFLOW
	static QString rst;
	if (i < 26)
		return QString(0x41 + i);
	else
		return getlinename(i / 26 - 1) + QString(0x41 + i % 26);
}
/******************************************************************************************************/
DWORD WINAPI FileProcess(LPVOID param)
{
	int decoder = xdec;
	QString tabletitle;
	QStringList qsl, title, linetitle, rowtitle;

	//select decoder
	switch (decoder)
	{
	case QETDECODER_DEFAULT:
		;
		break;
	case QETDECODER_UNICODE:
		xcontent = QString(xcontent.unicode());
		break;
	case QETDECODER_LOCAL8BIT:
		xcontent = xcontent.toLocal8Bit();
		break;
	case QETDECODER_LATIN1:
		xcontent = xcontent.toLatin1();
		break;
	default:
		break;
	}

	while (xcontent.back() == QChar('\n'))
		xcontent.chop(1);

	qsl = xcontent.split('\n', QString::SkipEmptyParts);
	size_t ssize = xcontent.length();
	int rmax = 0, temp = 0;

	for (int i = 0; i < qsl.size(); i++)
	{
		rmax = 0;
		for (int j = 0; j < qsl.at(i).length(); j++)
			rmax += (qsl.at(i).at(j) == ',');
		ntvrow = ntvrow > rmax ? ntvrow : rmax;
		//get tabletitle
		QStringList lline = qsl.at(i).split(',');
		int size = lline.size();
		temp = size;
		for (int j = 0; j < lline.size(); j++)
			temp -= isempty(lline.at(j));
		if (temp == 1)
		{
			title << qsl.at(i).split(',', QString::SkipEmptyParts).at(0);
			qsl.replace(i, qsl.at(i).split(',', QString::SkipEmptyParts).at(0));
		}
		lline = qsl.at(i).split(',');
		for (int j = 0; j < lline.length() - 1; j++)
			if (!isempty(lline.at(j)) && (isempty(lline.at(j + 1))))
				lline.replace(j + 1, lline.at(j));
		qsl.replace(i, lline.join(','));
	}
	ntvline = qsl.size() + 1;
	tabletitle = title.join(' ');
	title.clear();

	model->setColumnCount(ntvrow);
	model->setRowCount(ntvline);
	QStringList linename, rowname, qsltp;
	QString qstp;
	//QList<QStandardItem*> qsi;
	for (int i = 1; i < ntvline; i++)
		model->setHeaderData(i - 1, Qt::Vertical, QString::fromLocal8Bit("列") + QString::number(i));
	for (int i = 0; i < ntvrow; i++)
		model->setHeaderData(i, Qt::Horizontal, QString::fromLocal8Bit("行") + getlinename(i));

	QList<QStandardItem*> data;



	for (int i = 0; i < ntvline; i++)
	{
		for (int j = 0; j < ntvrow; j++)
		{
			//data<< new QStandardItem(
			model->setItem(i, j, new QStandardItem(QString::fromLocal8Bit("张三")));
		}
	}

	//model->setItem(0, 0, new QStandardItem("张三"));
	//model->setItem(0, 1, new QStandardItem("3"));
	//model->setItem(0, 2, new QStandardItem("男"));

	xtableView->setModel(model);

	//qDebug("%d,%d", ntvrow, ntvline);
	;
	xstatusBar->showMessage(QString::fromLocal8Bit("就绪"));
	return 0;
}
/******************************************************************************************************/
void QExcelTool::OpenFile()
{
	//reset global variables
	err = 0, ftype = -1, ntvline = 0, ntvrow = 0, xdec = 0;
	xfile = nullptr;
	xpath.clear();
	xcontent.clear();
	//open file
	xpath = QFileDialog::getOpenFileName(this, QString::fromLocal8Bit("打开表格文件"), ".", QString::fromLocal8Bit("表格文件(*.csv)")).replace("/", "\\");
	ui.textBrowser->setText(xpath);
	if (xpath.isEmpty())
	{
		ui.statusBar->showMessage(QString::fromLocal8Bit("错误：未选择文件"));
		err = QETERROR_NOFILESELECTED;
	}
	else
	{
		xfile = new QFile(xpath);
		xfile->open(QIODevice::ReadOnly | QIODevice::Text);
		if (!xfile->isOpen())//open file failed
		{
			ui.statusBar->showMessage(QString::fromLocal8Bit("错误：打开文件失败（%1)").arg(xfile->errorString()));
			err = QETERROR_FILEOPENFAIL;
		}
		else
		{
			QString fname = xpath.mid(xpath.lastIndexOf('\\') + 1);
			xstatusBar->showMessage(QString::fromLocal8Bit("信息：打开文件成功\t文件名：\"%1\"").arg(fname));
			QString fext = fname.mid(fname.lastIndexOf('.') + 1);
			ftype = (fext == "csv" ? QETFILETYPE_CSV : (fext == "xls" ? QETFILETYPE_XLS : (fext == "xlsx" ? QETFILETYPE_XLSX : -1)));
			if (ftype == -1)
			{
				QMessageBox::critical(this, "ERROR", "Unknown Error");
				QApplication::quit();
			}
			err = QETERROR_NOERROR;
			CreateThread(NULL, 0, ReadFile, NULL, NULL, NULL);
		}
	}
}
/******************************************************************************************************/
void QExcelTool::ChangeDec()
{
	xdec += 1;
	xdec %= QETDECODER_DECCOUNT;
	xqpb->setText(QString::fromLocal8Bit("切换编码(%1)").arg(xdec));
	if (isempty(xpath))
		return;
	CreateThread(NULL, 0, FileProcess, NULL, NULL, NULL);
}
/******************************************************************************************************/
//new thread
DWORD WINAPI ReadFile(LPVOID param)
{
	xstatusBar->showMessage(QString::fromLocal8Bit("信息：读取文件...\t文件名：\"%1\"").arg(xpath));
	xcontent = xfile->readAll();
	xfile->close();
	xstatusBar->showMessage(QString::fromLocal8Bit("就绪"));
	CreateThread(NULL, 0, FileProcess, NULL, NULL, NULL);
	return 0;
}
