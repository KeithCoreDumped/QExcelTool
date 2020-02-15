#include "QExcelTool.h"
#include <QFile>
#include <QMessageBox>
#include <QImage>
#include <QStandardItemModel>
#include <QFileDialog>
#include <QTextCodec>

int err = QETERROR_NOFILESELECTED, ntvrow = 0, ntvcolumn = 0, xdec = 0;
int fponstuff = 0, comboBoxlastIndex = 0, qtvedited = 0;
QStatusBar* xstatusBar = nullptr;
QFile* xfile = nullptr;
QTableView* xtableView = nullptr;
QComboBox* xcomboBox;
QString xpath, xcontent;
QByteArray xraw;
QStandardItemModel* model;
QTextCodec* xcodec;

typedef enum tagTextCodeType
{
	TextUnkonw = -1,
	TextANSI = 0,
	TextUTF8,
	TextUNICODE,
	TextUNICODE_BIG
}TextCodeType;
TextCodeType xftype = TextUnkonw;
/******************************************************************************************************/
inline bool isempty(QString src)
{
	return src.isNull() || src.isEmpty();
}
/******************************************************************************************************/
QExcelTool::QExcelTool(QWidget* parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	ui.statusBar->setSizeGripEnabled(false);
	ui.tableView->setShowGrid(true);
	xfile = nullptr;
	xtableView = ui.tableView;
	connect(ui.pushButton_1, SIGNAL(clicked()), this, SLOT(OpenFile()));
	connect(ui.comboBox, SIGNAL(activated(const QString&)), this, SLOT(oncomboBoxcurrentIndexChanged(const QString&)));

	xstatusBar = ui.statusBar;
	//ui.comboBox->setVisible(0);
	ui.statusBar->showMessage(QString::fromLocal8Bit("就绪"));
	ui.textBrowser->setToolTipDuration(1000);
	model = new QStandardItemModel();
	//QStringList labels = QString::fromLocal8Bit("1,2,3").simplified().split(",");
	//model->setHorizontalHeaderLabels(labels);
	//ui.tableView->setModel(model);
	//ui.tableView->show();
}
/******************************************************************************************************/
QString getcolumnname(unsigned int i)
{
	///STACK OVERFLOW
	static QString rst;
	if (i < 26)
		return QString(0x41 + i);
	else
		return getcolumnname(i / 26 - 1) + QString(0x41 + i % 26);
}
/******************************************************************************************************/
void QExcelTool::OpenFile()
{
	//reset global variables
	err = QETERROR_NOFILESELECTED, ntvrow = 0, ntvcolumn = 0, xdec = 0;
	xfile = nullptr;
	xcomboBox = ui.comboBox;
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
			//QString fname = xpath.mid(xpath.lastIndexOf('\\') + 1);
			//xstatusBar->showMessage(QString::fromLocal8Bit("信息：打开文件成功\t文件名：\"%1\"").arg(fname));
			//QString fext = fname.mid(fname.lastIndexOf('.') + 1);
			//ftype = (fext == "csv" ? QETFILETYPE_CSV : (fext == "xls" ? QETFILETYPE_XLS : (fext == "xlsx" ? QETFILETYPE_XLSX : -1)));
			//if (ftype == -1)
			//{
			//	QMessageBox::critical(this, "ERROR", "Unknown Error");
			//	QApplication::quit();
			//}
			err = QETERROR_NOERROR;
			CreateThread(NULL, 0, ReadFile, NULL, NULL, NULL);
		}
	}
}
/******************************************************************************************************/
void QExcelTool::oncomboBoxcurrentIndexChanged(const QString& arg1)
{
	if (fponstuff)
	{
		ui.comboBox->setCurrentIndex(comboBoxlastIndex);
		return;
	}
	if (QString::compare(arg1, "ANSI") == 0)
		xftype = TextANSI;
	else if (QString::compare(arg1, "UTF-8") == 0)
		xftype = TextUTF8;
	else if (QString::compare(arg1, "UTF-16LE") == 0)
		xftype = TextUNICODE;
	else if (QString::compare(arg1, "UTF-16BE") == 0)
		xftype = TextUNICODE_BIG;
	comboBoxlastIndex = xcomboBox->currentIndex();
	if ((!err))
		FileProcess(0);
		//CreateThread(NULL, 0, FileProcess, NULL, NULL, NULL);
}
/******************************************************************************************************/
//new thread
DWORD WINAPI ReadFile(LPVOID param)
{
	xstatusBar->showMessage(QString::fromLocal8Bit("信息：读取文件...\t文件名：\"%1\"").arg(xpath));
	int decoder = xdec;
	xraw = xfile->readAll();
	xfile->close();
	unsigned char headBuf[3] = { 0 };
	headBuf[0] = xraw.at(0);
	headBuf[1] = xraw.at(1);
	headBuf[2] = xraw.at(2);
	if (headBuf[0] == 0xEF && headBuf[1] == 0xBB && headBuf[2] == 0xBF)//utf8-bom 文件开头：FF BB BF,不带bom稍后解决
		xftype = TextUTF8;
	else if (headBuf[0] == 0xFF && headBuf[1] == 0xFE)//小端Unicode  文件开头：FF FE  intel x86架构自身是小端存储，可直接读取
		xftype = TextUNICODE;
	else if (headBuf[0] == 0xFE && headBuf[1] == 0xFF)//大端Unicode  文件开头：FE FF
		xftype = TextUNICODE_BIG;
	else
		xftype = TextANSI;    //ansi或者unf8 无bom

	CreateThread(NULL, 0, FileProcess, NULL, NULL, NULL);
	return 0;
}
/******************************************************************************************************/
DWORD WINAPI FileProcess(LPVOID param)
{
	xstatusBar->showMessage(QString::fromLocal8Bit("信息：正在处理..."));
	QString tabletitle;
	QStringList qsl, title, linetitle, rowtitle;
	model->clear();
	switch (xftype)
	{
	case TextUTF8:
		xcodec = QTextCodec::codecForName("UTF-8");
		break;
	case TextUNICODE:
		xcodec = QTextCodec::codecForName("UTF-16LE");
		break;
	case TextUNICODE_BIG:
		xcodec = QTextCodec::codecForName("UTF-16BE");
		break;
	case TextANSI:
		xcodec = QTextCodec::codecForName("GBK");
		break;
	default:
		break;
	}
	xcontent = xcodec->toUnicode(xraw);
	while (xcontent.back() == QChar('\n'))
		xcontent.chop(1);
	qsl = xcontent.split('\n', QString::SkipEmptyParts);
	size_t ssize = xcontent.length();
	int rmax = 0, temp = 0;
	ntvcolumn = 0, ntvrow = 0;
	for (int i = 0; i < qsl.size(); i++)
	{
		rmax = 0;
		for (int j = 0; j < qsl.at(i).length(); j++)
			rmax += (qsl.at(i).at(j) == ',');
		ntvcolumn = ntvcolumn > rmax ? ntvcolumn : rmax;
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
	ntvrow = qsl.size() + 1;
	tabletitle = title.join(' ');
	title.clear();
	ntvcolumn++;
	model->setColumnCount(ntvcolumn);
	model->setRowCount(ntvrow);
	QStringList qsltp;
	for (int i = 0; i < ntvrow; i++)
		model->setHeaderData(i, Qt::Vertical, QString::fromLocal8Bit("列") + QString::number(i + 1));
	for (int i = 0; i < ntvcolumn + 1; i++)
		model->setHeaderData(i, Qt::Horizontal, QString::fromLocal8Bit("行") + getcolumnname(i));

	for (int i = 0; i < ntvrow - 1; i++)
	{
		qsltp = qsl.at(i).split(',');
		temp = qsltp.size();
		for (int j = 0; j < temp; j++)
			model->setItem(i, j, new QStandardItem(qsltp.at(j)));
	}

	xtableView->setModel(model);
	xstatusBar->showMessage(QString::fromLocal8Bit("就绪"));
	return 0;
}
