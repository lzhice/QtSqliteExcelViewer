#include "excelengine.h"
#include "qt_windows.h"

CExcelEngine::CExcelEngine()
	: m_pExcel(NULL),
	m_pWorkbooks(NULL),
	m_pWorkbook(NULL),
	m_pWorksheet(NULL)
{
	init();
}

CExcelEngine::CExcelEngine(const QString& xlsFile)
	: m_sXlsFile(xlsFile),
	m_pExcel(NULL),
	m_pWorkbooks(NULL),
	m_pWorkbook(NULL),
	m_pWorksheet(NULL)
{
	init();
}

CExcelEngine::~CExcelEngine()
{
    if ( m_bIsOpen )
    {
        //����ǰ���ȱ������ݣ�Ȼ��ر�workbook
        close();
    }
    OleUninitialize();
}

void CExcelEngine::init()
{
	m_nRowCount    = 0;
	m_nColumnCount = 0;
	m_nStartRow    = 0;
	m_nStartColumn = 0;

	m_bIsOpen     = false;
	m_bIsValid    = false;
	m_bNewFile = false;

	m_pExcel = new QAxObject("Excel.Application", this);
	m_pWorkbooks = m_pExcel->querySubObject("Workbooks");

	HRESULT r = OleInitialize(0);
	if (r != S_OK && r != S_FALSE)
	{
		qDebug("Qt: Could not initialize OLE (error %x)", (unsigned int)r);
	}
}

/**
  *@brief ��sXlsFileָ����excel����
  *@return true : �򿪳ɹ�
  *        false: ��ʧ��
  */
bool CExcelEngine::open(quint32 nSheet, bool visible)
{
    if (m_bIsOpen)
    {
        close();
    }

	if (m_pExcel->isNull() || m_sXlsFile.isEmpty())
	{
		m_bIsOpen = false;
		return false;
	}

	m_pExcel->dynamicCall("SetVisible(bool)", visible);

    /*���ָ����ļ������ڣ�����Ҫ�½�һ��*/
    QFile file(m_sXlsFile);
    if (!file.exists())
    {
        m_bNewFile = true;
		m_pWorkbooks->dynamicCall("Add(void)");	//���һ���µĹ�����
		m_pWorkbook = m_pExcel->querySubObject("ActiveWorkBook");
    }
    else
    {
        m_bNewFile = false;
		m_pWorkbook = m_pWorkbooks->querySubObject("Open(const QString&)", m_sXlsFile);	//��xls��Ӧ�Ĺ�����
    }
	
	if (!m_pWorkbook->isNull())
	{
		QAxObject *pWorksheets = m_pWorkbook->querySubObject("WorkSheets");
		m_pWorksheet = pWorksheets->querySubObject("Item(int)", nSheet);	//��ָ����sheet
		if (!m_pWorksheet->isNull())
		{
			//�����Ѵ򿪣���ʼ��ȡ��Ӧ����
			QAxObject *usedrange = m_pWorksheet->querySubObject("UsedRange");		//��ȡ��sheet��ʹ�÷�Χ����
			if (!usedrange->isNull())
			{
				//��Ϊexcel���Դ��������������ݶ���һ���Ǵ�0,0��ʼ�����Ҫ��ȡ�������±�
				m_nStartRow    = usedrange->property("Row").toInt();		//��һ�е���ʼλ��
				m_nStartColumn = usedrange->property("Column").toInt();		//��һ�е���ʼλ��

				QAxObject *rows = usedrange->querySubObject("Rows");
				QAxObject *columns = usedrange->querySubObject("Columns");
				if (!rows->isNull() && !columns->isNull())
				{
					m_nRowCount    = rows->property("Count").toInt();			//��ȡ����
					m_nColumnCount = columns->property("Count").toInt();		//��ȡ����
				}
				m_bIsOpen  = true;
			}
		}
	}
	
    return m_bIsOpen;
}

/**
  *@brief Open()�����غ���
  */
bool CExcelEngine::open(const QString& xlsFile, quint32 nSheet, bool visible)
{
    m_sXlsFile = xlsFile;
    return open(nSheet, visible);
}

/**
  *@brief ���������ݣ�������д���ļ�
  */
void CExcelEngine::save()
{
    if (!m_pWorkbook->isNull())
    {
        if (!m_bNewFile)
        {
            m_pWorkbook->dynamicCall("Save()");
        }
        else /*������ĵ����½������ģ���ʹ�����ΪCOM�ӿ�*/
        {
            m_pWorkbook->dynamicCall("SaveAs(const QString&)", m_sXlsFile);
        }
    }
}

/**
  *@brief �ر�ǰ�ȱ������ݣ�Ȼ��رյ�ǰExcel COM���󣬲��ͷ��ڴ�
  */
void CExcelEngine::close()
{
    if (!m_pExcel->isNull() && !m_pWorkbook->isNull())
    {
        m_pWorkbook->dynamicCall("Close(bool)", true);
        m_pExcel->dynamicCall("Quit(void)");

        delete m_pExcel;
        m_pExcel = NULL;

        m_bIsOpen   = false;
        m_bNewFile	= false;
    }
}

/**
  *@brief ��QSqlTableModel�е����ݱ��浽excel��
  *@param model : ָ��GUI�е�QSqlTableModelָ��
  *@return ����ɹ���� true : �ɹ�
  *						false: ʧ��
  */
bool CExcelEngine::saveDataFromTable(QAbstractItemModel *pModel)
{
    if (!m_bIsOpen || NULL == pModel)
    {
        return false;
    }

	clearRangeCells(m_nStartRow, m_nStartColumn, m_nStartRow + m_nRowCount - 1, m_nStartColumn + m_nColumnCount - 1);

    int nRowCount		= pModel->rowCount();
    int nColumnCount	= pModel->columnCount();

    //��ȡ��ͷд����һ��
    for (int i = 0; i < nColumnCount; i++)
    {
		//qDebug() << model->headerData(i, Qt::Horizontal);
		setCellData(1, i+1, pModel->headerData(i, Qt::Horizontal));
    }

    //д����
    for (int i = 0; i < nRowCount; i++)
    {
        for (int j = 0; j < nColumnCount; j++)
        {
            if ( pModel->index(i, j) != QModelIndex() )
            {
				//qDebug() << model->index(i, j).data();
                setCellData(i+2, j+1, pModel->index(i, j).data());
            }
        }
    }

	setAllColumnAutoFit();
	setAllCellTextCenter();

    //����
    save();

    return true;
}

/**
  *@brief ��ָ����xls�ļ��а����ݵ��뵽tableWidget��
  *@param tableWidget : ִ��Ҫ���뵽��tablewidgetָ��
  *@return ����ɹ���� true : �ɹ�
  *                     false: ʧ��
  */
bool CExcelEngine::readDataToTable(QTableWidget *pTableWidget)
{
    if ( NULL == pTableWidget )
    {
        return false;
    }

    //�Ȱ�table���������
    pTableWidget->clear();
	pTableWidget->setColumnCount(0);
	pTableWidget->setRowCount(0);

    int rowcnt    = m_nStartRow + m_nRowCount;
    int columncnt = m_nStartColumn + m_nColumnCount;

    //��ȡexcel�еĵ�һ��������Ϊ��ͷ
    QStringList headerList;
    for (int n = 0; n < columncnt; n++ )
    {
        QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", m_nStartRow, n+m_nStartColumn);
        if (!cell->isNull())
        {
            headerList << cell->dynamicCall("Value2()").toString();
        }
    }
	pTableWidget->setColumnCount(m_nColumnCount);
    pTableWidget->setHorizontalHeaderLabels(headerList);

    //����������
    for (int i = m_nStartRow + 1, r = 0; i < rowcnt; ++i, ++r)  //��
    {
        pTableWidget->insertRow(r); //��������
        for (int j = m_nStartColumn, c = 0; j < columncnt; j++, c++)  //��
        {
            QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", i, j );//��ȡ��Ԫ��
            if (!cell->isNull())
            {
                pTableWidget->setItem(r,c,new QTableWidgetItem(cell->dynamicCall("Value2()").toString()));
            }
        }
    }

    return true;
}

bool CExcelEngine::readDataToTable(CTableModel *model)
{
	model->reset();

	QStringList strlstHeaders;
	//��ȡexcel�еĵ�һ��������Ϊ��ͷ
	for (int n = 0; n < m_nColumnCount; n++)
	{
		QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", m_nStartRow, n+m_nStartColumn);
		if (!cell->isNull())
		{
			QString strValue = cell->dynamicCall("Value2()").toString();
			strlstHeaders << strValue;
		}
	}
	model->setHorizontalHeaders(strlstHeaders);
	model->insertRows(0, m_nRowCount-1);
	
	//����������
	for (int i = 0; i < m_nRowCount-1; ++i)  //��
	{
		for (int j = 0; j < m_nColumnCount; ++j)  //��
		{
			QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", i+m_nStartRow+1, j+m_nStartColumn);//��ȡ��Ԫ��
			if (!cell->isNull())
			{
				QVariant value = cell->dynamicCall("Value2()");
				const QModelIndex& index = model->index(i, j);
				if (index.isValid())
				{
					model->setData(index, value);
				}
				else
				{
					qDebug() << "inValid index!";
				}
			}
			else
			{
				qDebug() << "error cell!";
			}
		}
	}

	return true;
}

/**
  *@brief ��ȡָ����Ԫ�������
  *@param row : ��Ԫ����к�
  *@param column : ��Ԫ����к�
  *@return [row,column]��Ԫ���Ӧ������
  */
QVariant CExcelEngine::getCellData(quint32 row, quint32 column)
{
    QVariant data;
    QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", row, column);//��ȡ��Ԫ�����
    if (!cell->isNull())
    {
        data = cell->property("Value");
    }

    return data;
}

/**
  *@brief �޸�ָ����Ԫ�������
  *@param row : ��Ԫ����к�
  *@param column : ��Ԫ��ָ�����к�
  *@param data : ��Ԫ��Ҫ�޸�Ϊ��������
  *@return �޸��Ƿ�ɹ� true : �ɹ�
  *						false: ʧ��
  */
bool CExcelEngine::setCellData(quint32 row, quint32 column, QVariant data)
{
    bool op = false;

    QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", row, column);//��ȡ��Ԫ�����
    if (!cell->isNull())
    {
        QString strData = data.toString();							//excel ��Ȼֻ�ܲ����ַ��������ͣ��������޷�����
        cell->dynamicCall("SetValue(const QVariant&)", strData);	//�޸ĵ�Ԫ�������
        op = true;
    }

    return op;
}

/**
  *@brief ��ճ�����֮�������
  */
void CExcelEngine::clear()
{
    m_sXlsFile     = "";
    m_nRowCount    = 0;
    m_nColumnCount = 0;
    m_nStartRow    = 0;
    m_nStartColumn = 0;
}

/**
  *@brief �ж�excel�Ƿ��ѱ���
  *@return true : �Ѵ�
  *        false: δ��
  */
bool CExcelEngine::isOpen()
{
    return m_bIsOpen;
}

/**
  *@brief �ж�excel COM�����Ƿ���óɹ���excel�Ƿ����
  *@return true : ����
  *        false: ������
  */
bool CExcelEngine::isValid()
{
    return m_bIsValid;
}

/**
  *@brief ��ȡexcel������
  */
quint32 CExcelEngine::rowCount()const
{
    return m_nRowCount;
}

/**
  *@brief ��ȡexcel������
  */
quint32 CExcelEngine::columnCount()const
{
    return m_nColumnCount;
}

/**
  *@brief ���excel��sheet��ָ����Ԫ�������
  */
void CExcelEngine::clearCell(int row, int column)
{
	if (!m_pWorksheet->isNull())
	{
		QString cell;
		cell.append(QChar(column - 1 + 'A'));
		cell.append(QString::number(row));

		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			range->dynamicCall("ClearContents()");
		}
	}
}

/**
  *@brief ���excel��sheet��ָ����Χ�����е�Ԫ�������
  */
void CExcelEngine::clearRangeCells(int rowFrom, int columnFrom, int rowTo, int columnTo)
{
	if (!m_pWorksheet->isNull())
	{
		QString cell;
		cell.append(QChar(columnFrom - 1 + 'A'));
		cell.append(QString::number(rowFrom));
		cell.append(":");
		cell.append(QChar(columnTo - 1 + 'A'));
		cell.append(QString::number(rowTo));

		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			range->dynamicCall("ClearContents()");
		}
	}
}

void CExcelEngine::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
	if (!m_pWorksheet->isNull())
	{
		QAxObject *usedRange = m_pWorksheet->querySubObject("UsedRange");
		if (usedRange)
		{
			*topLeftRow = usedRange->property("Row").toInt();
			*topLeftColumn = usedRange->property("Column").toInt();

			QAxObject *rows = usedRange->querySubObject("Rows");
			QAxObject *columns = usedRange->querySubObject("Columns");

			if (rows && columns)
			{
				*bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;
				*bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
			}
		}
	}
}

void CExcelEngine::setColumnWidth(int column, int width)
{
	if (!m_pWorksheet->isNull())
	{
		QString columnName;
		columnName.append(QChar(column - 1 + 'A'));
		columnName.append(":");
		columnName.append(QChar(column - 1 + 'A'));

		QAxObject *cols = m_pWorksheet->querySubObject("Columns(const QString&)", columnName);
		if (cols)
		{
			cols->setProperty("ColumnWidth", width);	
		}			
	}
}

void CExcelEngine::setRowHeight(int row, int height)
{
	if (!m_pWorksheet->isNull())
	{
		QString rowsName;
		rowsName.append(QString::number(row));
		rowsName.append(":");
		rowsName.append(QString::number(row));

		QAxObject *rows = m_pWorksheet->querySubObject("Rows(const QString &)", rowsName);
		if (rows)
		{
			rows->setProperty("RowHeight", height);
		}
	}
}

void CExcelEngine::setCellTextWrap(int row, int column, bool isWrap)
{
	if (!m_pWorksheet->isNull())
	{
		QString cell;
		cell.append(QChar(column - 1 + 'A'));
		cell.append(QString::number(row));

		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			range->setProperty("WrapText", isWrap);
		}
	}
}

void CExcelEngine::setAllColumnAutoFit()
{
	QAxObject *usedrange = m_pWorksheet->querySubObject("UsedRange");		//��ȡ��sheet��ʹ�÷�Χ����
	if (usedrange)
	{
		QAxObject *columns = usedrange->querySubObject("Columns");
		if (columns)
		{
			columns->dynamicCall("AutoFit()");
		}
	}
}

void CExcelEngine::setAutoFitColumn(int columnFrom, int columnTo)
{
	if (!m_pWorksheet->isNull())
	{
		QString strRange;
		strRange.append(QChar(columnFrom - 1 + 'A'));
		strRange.append(":");
		strRange.append(QChar(columnTo - 1 + 'A'));

		QAxObject *columns = m_pWorksheet->querySubObject("Columns(const QString &)", strRange);
		if (columns)
		{
			columns->dynamicCall("AutoFit()");
		}
	}
}

void CExcelEngine::setCellTextCenter(int row, int column)
{
	if (!m_pWorksheet->isNull())
	{
		QString cell;
		cell.append(QChar(column - 1 + 'A'));
		cell.append(QString::number(row));

		setCellTextCenter(cell);
	}
}

void CExcelEngine::setCellTextCenter(const QString &cell)
{
	if (!m_pWorksheet->isNull())
	{
		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			range->setProperty("HorizontalAlignment", -4108);//xlCenter
		}
	}
}

void CExcelEngine::setAllCellTextCenter()
{
	if (!m_pWorksheet->isNull())
	{
		QAxObject *usedrange = m_pWorksheet->querySubObject("UsedRange");
		if (usedrange)
		{
			usedrange->setProperty("HorizontalAlignment", -4108);//xlCenter
		}
	}
}

void CExcelEngine::setCellTextWrap(const QString &cell, bool isWrap)
{
	if (!m_pWorksheet->isNull())
	{
		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			range->setProperty("WrapText", isWrap);
		}
	}
}

void CExcelEngine::setCellFontBold(const QString &cell, bool isBold)
{
	if (!m_pWorksheet->isNull())
	{
		QAxObject *range = m_pWorksheet->querySubObject("Range(const QString&)", cell);
		if (range)
		{
			QAxObject *font = range->querySubObject("Font");
			font->setProperty("Bold", isBold);
		}
	}
}