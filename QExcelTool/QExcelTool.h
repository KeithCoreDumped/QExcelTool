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

extern int err, ftype, ntvline, ntvrow;
extern QStatusBar* xstatusBar;
extern QFile* xfile;
extern std::string content;
extern QTableView* xtableView;


class QExcelTool : public QMainWindow
{
	Q_OBJECT
public slots:
	void openxlsfile();

public:
	QExcelTool(QWidget* parent = Q_NULLPTR);

private:
	Ui::QExcelToolClass ui;
	QFile* fxls;
	QString fpath;
};

DWORD WINAPI ReadFile(LPVOID param);
