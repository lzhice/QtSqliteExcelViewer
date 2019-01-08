#include "sqlexcelviewer.h"
#include <QSqlDriver>
#include <QSqlQuery>
#include <QSqlError>
#include <QStandardPaths>
#include <QFile>
#include <QFileDialog>
#include <QDebug>
#include <QMessageBox>
#include <QDesktopWidget>
#include "PwdDlg.h"

#define QTSQLEXCEL_DB_CONN QStringLiteral("QTSQLEXCEL_DB_CONNECTION")
#define SMALLSIZE 260
#define LARGESIZE 620

CSqlExcelViewer::CSqlExcelViewer(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	setWindowTitle(QString::fromStdWString(L"Qt Sqlite�������ݿ�鿴��"));
	setFixedSize(695, SMALLSIZE);
	initCtrl();

	connect(ui.act_open, SIGNAL(triggered()), this, SLOT(onSelectFile()));
	connect(ui.act_Close, SIGNAL(triggered()), this, SLOT(onCloseFile()));
	connect(ui.actionImport, SIGNAL(triggered()), this, SLOT(onImportFrExcel()));
	connect(ui.actionExport, SIGNAL(triggered()), this, SLOT(onExportToExcel()));
	connect(ui.btnShowAll, SIGNAL(clicked()), this, SLOT(onRefreshTable()));
	connect(ui.btnSubmit, SIGNAL(clicked()), this, SLOT(onSubmitChanges()));
	connect(ui.btnRevert, SIGNAL(clicked()), this, SLOT(onRevertChanges()));
	connect(ui.editKey, SIGNAL(textChanged(const QString &)), this, SLOT(onKeyChanged(const QString &)));
	connect(ui.cmbTables, SIGNAL(currentIndexChanged(const QString &)), this, SLOT(onTableChanged(const QString &)));
	connect(ui.btnQuery, SIGNAL(clicked()), this, SLOT(onQuery()));
	connect(ui.btnAdd, SIGNAL(clicked()), this, SLOT(onAddRow()));
	connect(ui.btnDelete, SIGNAL(clicked()), this, SLOT(onRemoveRow()));
}

CSqlExcelViewer::~CSqlExcelViewer()
{
	removeDB();
}

void CSqlExcelViewer::initCtrl()
{
	addDB();

	//m_pSqlTableModel�������db�ļ��е�����
	m_pSqlTableModel = new QSqlTableModel(this, m_db);
	m_pSqlTableModel->setEditStrategy(QSqlTableModel::OnManualSubmit); //���ñ������Ϊ�ֶ��ύ

	//m_pTableModel�������Excel�е�����
	m_pTableModel = new CTableModel(this);

	m_proxyModel = new QSortFilterProxyModel(this);
	m_proxyModel->setSourceModel(m_pSqlTableModel);

	///TableView������ʾ����
	m_pTableView = new QTableView(this);
	m_pTableView->setGeometry(5, 270, 685, 350);
	m_pTableView->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
	m_pTableView->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
	m_pTableView->setSortingEnabled(true);
	m_pTableView->sortByColumn(0, Qt::AscendingOrder);
	m_pTableView->setModel(m_proxyModel);
	m_pTableView->hide();

	ui.btnShowAll->setEnabled(false);
	ui.cmbRecord->setEditable(false);
	ui.cmbTables->setEditable(false);
}

void CSqlExcelViewer::addDB()
{
	if(QSqlDatabase::isDriverAvailable("QSQLITECIPHER"))
	{
		m_db = QSqlDatabase::addDatabase("QSQLITECIPHER", QTSQLEXCEL_DB_CONN);
	}
	else
	{
		qWarning("Can not find sqlitecipher.dll.");
	}
	Q_ASSERT(m_db.isValid());
}

void CSqlExcelViewer::removeDB()
{
	if (m_db.isValid())
	{
		if (m_db.isOpen())
		{
			m_db.close();
		}
		m_db = QSqlDatabase();
		QSqlDatabase::removeDatabase(QTSQLEXCEL_DB_CONN);

		m_strDBFile.clear();
		m_strTableName.clear();
		qDebug() << "valid: " << m_db.isValid();
	}
}

bool CSqlExcelViewer::connectToDatabase(QString strDBFile)
{
	if (m_db.isValid() && m_strDBFile != strDBFile)
	{
		m_db.setDatabaseName(strDBFile);
		if (!m_db.open()) //������
		{
			CPwdDlg pwdDlg(this);
			if (QDialog::Accepted == pwdDlg.exec())
			{
				if (!m_db.open("", pwdDlg.getPassword()))
				{
					qDebug() << m_db.lastError().text();
					QMessageBox::information(this, "Error", "Error Password!");
					return false;
				}
			}
		}

		m_strDBFile = strDBFile;
		synTablesToCombox();
	}
	return true;
}

void CSqlExcelViewer::setTableInfoVisible(bool bShow)
{
	m_pTableView->hide();
	if (bShow)
	{
		if (!m_strTableName.isEmpty())
		{
			if (height() != LARGESIZE)
			{
				setFixedSize(width(), LARGESIZE);
			}

			if (m_pSqlTableModel->tableName() != m_strTableName)
			{
				m_pSqlTableModel->setTable(m_strTableName);
				synTableFieldsToCombox();
			}

			m_pSqlTableModel->select();
			m_pSqlTableModel->sort(1, Qt::AscendingOrder);
			m_pTableView->resizeColumnsToContents();
			m_pTableView->show();

			ui.btnShowAll->setEnabled(true);
			ui.btnQuery->setEnabled(true);
		}
	}
	else
	{
		if (height() != SMALLSIZE)
		{
			setFixedSize(width(), SMALLSIZE);
		}
	}
	QRect rtScreen = QApplication::desktop()->screenGeometry();
	move(rtScreen.width()/2 - width()/2, rtScreen.height()/2 - height()/2);
}

void CSqlExcelViewer::synTablesToCombox()
{
	ui.cmbTables->clear();
	const QStringList& lstTables = getAllTables();

	if (!lstTables.isEmpty())
	{
		ui.cmbTables->addItems(lstTables);
		//Ĭ�϶���һ�ű�
		m_strTableName = lstTables[0];
	}
	else
	{
		QMessageBox::information(this, "Error", "There has no table in DB or the DB is not created by QSQLITECIPHER!");
	}
}

void CSqlExcelViewer::synTableFieldsToCombox()
{
	ui.cmbRecord->clear();

	QStringList lstHorizontalHead;
	for (int i = 0; i < m_pSqlTableModel->columnCount(); ++i)
	{
		lstHorizontalHead << m_pSqlTableModel->headerData(i, Qt::Horizontal).toString();
	}
	ui.cmbRecord->addItems(lstHorizontalHead);
}

QStringList CSqlExcelViewer::getTableFieldNames(QString strTableName, const QSqlDatabase &db)
{
	QStringList lstHorizontalHead;

	QSqlQuery query(db);
	QString strSql = QString("PRAGMA table_info([%1])").arg(strTableName);
	if (query.exec(strSql))
	{
		while (query.next())
		{
			qDebug() << query.value(0);
			lstHorizontalHead << query.value(1).toString();
		}
	}
	return lstHorizontalHead;
}

QVector<QSqlRecord> CSqlExcelViewer::getTableRecords(QString strTableName, const QSqlDatabase &db)
{
	QVector<QSqlRecord> vecRecords;

	QSqlQuery query(db);
	QString strSql = QString("select * from %1").arg(strTableName);
	if (query.exec(strSql))
	{
		if (m_db.driver()->hasFeature(QSqlDriver::QuerySize))
		{
			qDebug() << "query.size(): " << query.size();
		}
		
		while (query.next())
		{
			vecRecords << query.record();

			int nFieldCount = query.record().count();
			for (int i = 0; i < nFieldCount; ++i)
			{
				QSqlField field = query.record().field(i);
				if (field.isValid())
				{
					qDebug() << "QSqlField: " << field;
				}
			}
		}
	}
	return vecRecords;
}

QStringList CSqlExcelViewer::getAllTables()
{
	QStringList lstTableNames;
	if (isConnecting())
	{
		lstTableNames = m_db.tables();
	}
	return lstTableNames;
}

bool CSqlExcelViewer::isConnecting()
{
	return (m_db.isValid() && m_db.isOpen());
}

//////////////////////////////////////////////////////////////////////////
void CSqlExcelViewer::onSelectFile()
{
	m_pTableView->hide();
	QString strDBFile = QFileDialog::getOpenFileName(this, QString::fromStdWString(L"ѡ��DB�ļ�"), 
		QDir::currentPath(), QString::fromStdWString(L"DB�ļ�(*.db *.db3)"));
	if (!strDBFile.isEmpty() && !strDBFile.isNull())
	{
		QFileInfo info(strDBFile);
		if (info.isFile())
		{
			if (connectToDatabase(strDBFile))
			{
				m_proxyModel->setSourceModel(m_pSqlTableModel);
				setTableInfoVisible();
			}
		}
	}
	else
	{
		m_pTableView->show();
	}
}

void CSqlExcelViewer::onCloseFile()
{
	m_strDBFile.clear();
	m_strTableName.clear();
	
	ui.btnShowAll->setEnabled(false);
	ui.btnQuery->setEnabled(false);
	ui.cmbTables->clear();
	ui.cmbRecord->clear();

	setTableInfoVisible(false);
}

void CSqlExcelViewer::onRefreshTable()
{
	m_pSqlTableModel->setFilter("");
	m_pSqlTableModel->select();
}

void CSqlExcelViewer::onSubmitChanges()
{
	if (isConnecting())
	{
		m_db.transaction();		//��ʼ�������
		if (m_pSqlTableModel->submitAll())
		{
			m_db.commit();		//�ύ
		}
		else
		{
			m_db.rollback();	//�ع�
			qWarning(m_pSqlTableModel->lastError().text().toStdString().c_str());
		}
	}
}

void CSqlExcelViewer::onRevertChanges()
{
	if (isConnecting())
	{
		m_pSqlTableModel->revertAll();
	}
}

void CSqlExcelViewer::onKeyChanged(const QString &strKey)
{
	if (!strKey.isEmpty())
	{
		ui.btnQuery->setEnabled(true);
	}
}

void CSqlExcelViewer::onTableChanged(const QString &strTableName)
{
	if (!strTableName.isEmpty() && strTableName != m_strTableName)
	{
		m_strTableName = strTableName;
		setTableInfoVisible();
	}
}

void CSqlExcelViewer::onQuery()
{
	QString strKey = ui.editKey->text();
	QString strCmbText = ui.cmbRecord->currentText();
	if (!strCmbText.isEmpty())
	{
		if (!strKey.isEmpty())
		{
			QString strSql = QString("%2 = \'%3\'").arg(strCmbText).arg(strKey);
			m_pSqlTableModel->setFilter(strSql);
		}
		else
		{
			m_pSqlTableModel->setFilter("");
		}
		m_pSqlTableModel->select();
		m_pTableView->show();
	}
}

void CSqlExcelViewer::onAddRow()
{
	int rowNum = m_pSqlTableModel->rowCount();		//��ñ������
	m_pSqlTableModel->insertRow(rowNum);			//���һ��
}

void CSqlExcelViewer::onRemoveRow()
{
	int curRow = m_pTableView->currentIndex().row();
	if (curRow >= 0 && curRow < m_pSqlTableModel->rowCount())
	{
		m_pSqlTableModel->removeRow(curRow);

		int nResult = QMessageBox::warning(this, tr("Remove"), tr("Do you want to remove current row ?"), QMessageBox::Yes, QMessageBox::No); 
		if(nResult == QMessageBox::No)
		{
			m_pSqlTableModel->revertAll(); //�����ɾ��������
		}
		else
		{
			m_pSqlTableModel->submitAll(); //�����ύ�������ݿ���ɾ������
		}
	}
}

void CSqlExcelViewer::onExportToExcel()
{
	QString desktop_path = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
	QString strFile = QFileDialog::getSaveFileName(this, tr("Save File"), desktop_path, tr("Excel File (*.xlsx)"));
	strFile = QDir::toNativeSeparators(strFile);
	if (!strFile.isNull() && !strFile.isEmpty())
	{
		exportToExcel(strFile);
	}
}

void CSqlExcelViewer::onImportFrExcel()
{
	QString desktop_path = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
	QString strFile = QFileDialog::getOpenFileName(this, QString::fromStdWString(L"ѡ��excel�ļ�"), desktop_path, tr("Excel File (*.xlsx)"));
	strFile = QDir::toNativeSeparators(strFile);
	if (!strFile.isNull() && !strFile.isEmpty() && QFile::exists(strFile))
	{
		importFromExcel(strFile);
	}
}

void CSqlExcelViewer::exportToExcel(QString strFilePath)
{
	CExcelEngine excel(strFilePath);
	if (excel.open(1U, true))
	{
		if (excel.saveDataFromTable(m_pSqlTableModel))
		{
			qDebug() << "Export Success";
		}
		else
		{
			qDebug() << "Export Failed";
		}
		excel.close();
	}
}

void CSqlExcelViewer::importFromExcel(QString strFilePath)
{
	CExcelEngine excel(strFilePath);
	if (excel.open())
	{
		if (excel.readDataToTable(m_pTableModel))
		{
			qDebug() << "Import Success";
		}
		else
		{
			qDebug() << "Import Failed";
		}
		excel.close();

		if (height() != LARGESIZE)
		{
			setFixedSize(width(), LARGESIZE);
			QRect rtScreen = QApplication::desktop()->screenGeometry();
			move(rtScreen.width()/2 - width()/2, rtScreen.height()/2 - height()/2);
		}

		m_proxyModel->setSourceModel(m_pTableModel);
		m_pTableView->resizeColumnsToContents();
		m_pTableView->show();
	}
}

#if 0
void QtSqlTest::onAddRecordToTable()
{
	if (m_db.isValid() && m_db.open("", m_strDBPwd))
	{
		QString strNewRecordName = ui.editRecord->text();
		QString strNewRecordType = ui.editType->text();
		if ("" != strNewRecordName && "" != strNewRecordType)
		{
			QSqlQuery query(m_db);
			QString strSql = QString("alter table %1 add %2 %3 default ''").arg(m_strTableName).arg(strNewRecordName).arg(strNewRecordType);
			query.prepare(strSql);
			if(!query.exec())
			{
				qDebug() << query.lastError();
			}
		}

		m_db.close();
	}
}

void QtSqlTest::onDeleteRecordFromTable()
{
	if (m_db.isValid() && m_db.open("", m_strDBPwd))
	{
		QString strRecordName = ui.editRecord->text();
		if ("" != strRecordName)
		{
			QSqlQuery query(m_db);
			QString strSql = QString("alter table %1 drop column %2").arg(m_strTableName).arg(strRecordName);
			query.prepare(strSql);
			if(!query.exec())
			{
				qDebug() << query.lastError();
			}
		}

		m_db.close();
	}
}

void QtSqlTest::insertRecordToTable(QString strTableName)
{
	//��ѯ����������¼
	int count = 0;
	QString select_count_sql = QString("select count(*) from %1").arg(TABLE_NAME);
	QSqlQuery sql_query(m_db);
	if (sql_query.exec(select_count_sql))
	{
		while(sql_query.next())
		{
			count = sql_query.value(0).toInt();
		}
	}
	count++;

	QVariantList lstId;
	for (int i = count; i < count + 5; ++i)
	{
		lstId << i;
	}

	QStringList lstName;
	lstName << "tom" << "jack" << "Lisa" << "curry" << "sherry";

	QVariantList lstAge;
	lstAge << 15 << 16 << 17 << 18 << 19;

	QString insert_sql = QString("insert into %1 values (?, ?, ?)").arg(TABLE_NAME);
	sql_query.prepare(insert_sql);

	//��������
	sql_query.addBindValue(lstId);
	sql_query.addBindValue(lstName);
	sql_query.addBindValue(lstAge);
	if(!sql_query.execBatch())
	{
		qDebug() << sql_query.lastError();
	}
}

void QtSqlTest::updateRecordFromTable(QString strTableName)
{
	//��������
	QString update_sql = QString("update %1 set name = :name where id = :id").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(update_sql);
	sql_query.bindValue(":name", "jack");
	sql_query.bindValue(":id", 1);
	if(!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
}

void QtSqlTest::deleteRecordFromTable(QString strTableName)
{
	//ɾ��id����һ����¼
	QString delete_sql = QString("delete from %1 where id = (select max(id) from %1)").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(delete_sql);
	if(!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
}

void QtSqlTest::selectRecordFromTable(QString strTableName)
{
	//��ѯ��������
	QString select_all_sql = QString("select * from %1").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(select_all_sql);
	if(!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
	else
	{
		while(sql_query.next())
		{
			int id			= sql_query.value(0).toInt();
			QString name	= sql_query.value(1).toString();
			int age			= sql_query.value(2).toInt();

			qDebug() << QString("id:%1  name:%2  age:%3").arg(id).arg(name).arg(age);
		}
	}
}

void QtSqlTest::clearTableRecord(QString strTableName)
{
	//��ձ�
	QString clear_sql = QString("delete from %1").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(clear_sql);
	if(!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
}

void QtSqlTest::deldeteTable(QString strTableName)
{
	//ɾ����
	QString deleteTable_sql = QString("DROP TABLE %1").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(deleteTable_sql);
	if (!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
}

void QtSqlTest::createTable(QString strTableName)
{
	QString create_sql = QString("create table %1 (id int primary key, name varchar(30), age int)").arg(strTableName);
	QSqlQuery sql_query(m_db);
	sql_query.prepare(create_sql);
	if(!sql_query.exec())
	{
		qDebug() << sql_query.lastError();
	}
}
#endif