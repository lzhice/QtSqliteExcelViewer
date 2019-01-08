#include "sqlexcelviewer.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	a.setWindowIcon(QIcon("./resources/database_viewer.ico"));
	CSqlExcelViewer w;
	w.show();
	return a.exec();
}