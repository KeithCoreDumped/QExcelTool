#pragma once

#include <QtWidgets/QMainWindow>
#include <QFile>
#include <string>
#include <Windows.h>
#include "ui_QExcelTool.h"
#include "QtXlsxWriter\\xlsxdocument.h"
#include "QtXlsxWriter\\xlsxformat.h"
#include "QtXlsxWriter\\xlsxcellrange.h"
#include "QtXlsxWriter\\xlsxworksheet.h"

#define QETERROR_NOERROR		0
#define QETERROR_NOFILESELECTED	1
#define QETERROR_FILEOPENFAIL	2

#define QETFILETYPE_CSV			0
#define QETFILETYPE_XLS			1
#define QETFILETYPE_XLSX		2

#define QETDECODER_DEFAULT		0
#define QETDECODER_UNICODE		1
#define QETDECODER_LOCAL8BIT	2
#define QETDECODER_LATIN1		3
#define QETDECODER_DECCOUNT		4

extern int err, ntvrow, ntvcolumn;
extern QStatusBar* xstatusBar;
extern QFile* xfile;
extern QString xpath;
//extern std::string content;
extern QTableView* xtableView;



class QExcelTool : public QMainWindow
{
	Q_OBJECT


public slots:
	void OpenFile();
	void oncomboBoxcurrentIndexChanged(const QString& arg1);

public:
	QExcelTool(QWidget* parent = Q_NULLPTR);
private:
	Ui::QExcelToolClass ui;
};

DWORD WINAPI ReadFile(LPVOID param);
DWORD WINAPI FileProcess(LPVOID param);
