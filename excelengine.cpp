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
        //析构前，先保存数据，然后关闭workbook
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
  *@brief 打开sXlsFile指定的excel报表
  *@return true : 打开成功
  *        false: 打开失败
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

    /*如果指向的文件不存在，则需要新建一个*/
    QFile file(m_sXlsFile);
    if (!file.exists())
    {
        m_bNewFile = true;
		m_pWorkbooks->dynamicCall("Add(void)");	//添加一个新的工作薄
		m_pWorkbook = m_pExcel->querySubObject("ActiveWorkBook");
    }
    else
    {
        m_bNewFile = false;
		m_pWorkbook = m_pWorkbooks->querySubObject("Open(const QString&)", m_sXlsFile);	//打开xls对应的工作簿
    }
	
	if (!m_pWorkbook->isNull())
	{
		QAxObject *pWorksheets = m_pWorkbook->querySubObject("WorkSheets");
		m_pWorksheet = pWorksheets->querySubObject("Item(int)", nSheet);	//打开指定的sheet
		if (!m_pWorksheet->isNull())
		{
			//至此已打开，开始获取相应属性
			QAxObject *usedrange = m_pWorksheet->querySubObject("UsedRange");		//获取该sheet的使用范围对象
			if (!usedrange->isNull())
			{
				//因为excel可以从任意行列填数据而不一定是从0,0开始，因此要获取首行列下标
				m_nStartRow    = usedrange->property("Row").toInt();		//第一行的起始位置
				m_nStartColumn = usedrange->property("Column").toInt();		//第一列的起始位置

				QAxObject *rows = usedrange->querySubObject("Rows");
				QAxObject *columns = usedrange->querySubObject("Columns");
				if (!rows->isNull() && !columns->isNull())
				{
					m_nRowCount    = rows->property("Count").toInt();			//获取行数
					m_nColumnCount = columns->property("Count").toInt();		//获取列数
				}
				m_bIsOpen  = true;
			}
		}
	}
	
    return m_bIsOpen;
}

/**
  *@brief Open()的重载函数
  */
bool CExcelEngine::open(const QString& xlsFile, quint32 nSheet, bool visible)
{
    m_sXlsFile = xlsFile;
    return open(nSheet, visible);
}

/**
  *@brief 保存表格数据，把数据写入文件
  */
void CExcelEngine::save()
{
    if (!m_pWorkbook->isNull())
    {
        if (!m_bNewFile)
        {
            m_pWorkbook->dynamicCall("Save()");
        }
        else /*如果该文档是新建出来的，则使用另存为COM接口*/
        {
            m_pWorkbook->dynamicCall("SaveAs(const QString&)", m_sXlsFile);
        }
    }
}

/**
  *@brief 关闭前先保存数据，然后关闭当前Excel COM对象，并释放内存
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
  *@brief 把QSqlTableModel中的数据保存到excel中
  *@param model : 指向GUI中的QSqlTableModel指针
  *@return 保存成功与否 true : 成功
  *						false: 失败
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

    //获取表头写做第一行
    for (int i = 0; i < nColumnCount; i++)
    {
		//qDebug() << model->headerData(i, Qt::Horizontal);
		setCellData(1, i+1, pModel->headerData(i, Qt::Horizontal));
    }

    //写数据
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

    //保存
    save();

    return true;
}

/**
  *@brief 从指定的xls文件中把数据导入到tableWidget中
  *@param tableWidget : 执行要导入到的tablewidget指针
  *@return 导入成功与否 true : 成功
  *                     false: 失败
  */
bool CExcelEngine::readDataToTable(QTableWidget *pTableWidget)
{
    if ( NULL == pTableWidget )
    {
        return false;
    }

    //先把table的内容清空
    pTableWidget->clear();
	pTableWidget->setColumnCount(0);
	pTableWidget->setRowCount(0);

    int rowcnt    = m_nStartRow + m_nRowCount;
    int columncnt = m_nStartColumn + m_nColumnCount;

    //获取excel中的第一行数据作为表头
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

    //插入新数据
    for (int i = m_nStartRow + 1, r = 0; i < rowcnt; ++i, ++r)  //行
    {
        pTableWidget->insertRow(r); //插入新行
        for (int j = m_nStartColumn, c = 0; j < columncnt; j++, c++)  //列
        {
            QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", i, j );//获取单元格
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
	//获取excel中的第一行数据作为表头
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
	
	//插入新数据
	for (int i = 0; i < m_nRowCount-1; ++i)  //行
	{
		for (int j = 0; j < m_nColumnCount; ++j)  //列
		{
			QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", i+m_nStartRow+1, j+m_nStartColumn);//获取单元格
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
  *@brief 获取指定单元格的数据
  *@param row : 单元格的行号
  *@param column : 单元格的列号
  *@return [row,column]单元格对应的数据
  */
QVariant CExcelEngine::getCellData(quint32 row, quint32 column)
{
    QVariant data;
    QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", row, column);//获取单元格对象
    if (!cell->isNull())
    {
        data = cell->property("Value");
    }

    return data;
}

/**
  *@brief 修改指定单元格的数据
  *@param row : 单元格的行号
  *@param column : 单元格指定的列号
  *@param data : 单元格要修改为的新数据
  *@return 修改是否成功 true : 成功
  *						false: 失败
  */
bool CExcelEngine::setCellData(quint32 row, quint32 column, QVariant data)
{
    bool op = false;

    QAxObject *cell = m_pWorksheet->querySubObject("Cells(int,int)", row, column);//获取单元格对象
    if (!cell->isNull())
    {
        QString strData = data.toString();							//excel 居然只能插入字符串和整型，浮点型无法插入
        cell->dynamicCall("SetValue(const QVariant&)", strData);	//修改单元格的数据
        op = true;
    }

    return op;
}

/**
  *@brief 清空除报表之外的数据
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
  *@brief 判断excel是否已被打开
  *@return true : 已打开
  *        false: 未打开
  */
bool CExcelEngine::isOpen()
{
    return m_bIsOpen;
}

/**
  *@brief 判断excel COM对象是否调用成功，excel是否可用
  *@return true : 可用
  *        false: 不可用
  */
bool CExcelEngine::isValid()
{
    return m_bIsValid;
}

/**
  *@brief 获取excel的行数
  */
quint32 CExcelEngine::rowCount()const
{
    return m_nRowCount;
}

/**
  *@brief 获取excel的列数
  */
quint32 CExcelEngine::columnCount()const
{
    return m_nColumnCount;
}

/**
  *@brief 清除excel的sheet表指定单元格的内容
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
  *@brief 清除excel的sheet表指定范围的所有单元格的内容
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
	QAxObject *usedrange = m_pWorksheet->querySubObject("UsedRange");		//获取该sheet的使用范围对象
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