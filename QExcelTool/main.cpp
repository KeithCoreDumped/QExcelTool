#include "QExcelTool.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	QExcelTool w;
	w.show();
	return a.exec();
}
