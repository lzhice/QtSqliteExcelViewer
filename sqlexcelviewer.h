#ifndef QTSQLTEST_H
#define QTSQLTEST_H

#include <QtWidgets/QMainWindow>
#include "ui_qtsqltest.h"
#include <QTableView>
#include <QSortFilterProxyModel>
#include <QSqlTableModel>
#include <QSqlDatabase>
#include <QSqlRecord>
#include <QSqlField>

#include "excelengine.h"
#include "tablemodel.h"

class CSqlExcelViewer : public QMainWindow
{
	Q_OBJECT

public:
	CSqlExcelViewer(QWidget *parent = 0);
	~CSqlExcelViewer();

	//获取表中所有字段名
	QStringList getTableFieldNames(QString strTableName, const QSqlDatabase &db);
	QVector<QSqlRecord> getTableRecords(QString strTableName, const QSqlDatabase &db);

private:
	void initCtrl();
	void addDB();
	void removeDB();

	bool connectToDatabase(QString strDBFile);
	bool isConnecting();

	void synTablesToCombox();
	void synTableFieldsToCombox();
	QStringList getAllTables();

	void setTableInfoVisible(bool bShow = true);

	void exportToExcel(QString strFilePath);
	void importFromExcel(QString strFilePath);

	/*
	void createTable(QString strTableName);
	void insertRecordToTable(QString strTableName);
	void updateRecordFromTable(QString strTableName);
	void selectRecordFromTable(QString strTableName);
	void deleteRecordFromTable(QString strTableName);
	void clearTableRecord(QString strTableName);
	void deldeteTable(QString strTableName);*/

	private Q_SLOTS:
		void onSelectFile();
		void onCloseFile();
		void onRefreshTable();
		void onSubmitChanges();
		void onRevertChanges();

		void onQuery();
		void onAddRow();
		void onRemoveRow();

		void onKeyChanged(const QString &);
		void onTableChanged(const QString &);

		void onExportToExcel();
		void onImportFrExcel();

		//void onAddRecordToTable();
		//void onDeleteRecordFromTable();

private:
	Ui::QtSqlTestClass ui;

	QSqlDatabase m_db;
	QTableView *m_pTableView;
	QSortFilterProxyModel *m_proxyModel;
	QSqlTableModel *m_pSqlTableModel;
	CTableModel *m_pTableModel;

	QString m_strDBFile;
	QString m_strTableName;
};

#endif // QTSQLTEST_H
