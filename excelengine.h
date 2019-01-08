#ifndef EXCELENGINE_H
#define EXCELENGINE_H

#include <QObject>
#include <QFile>
#include <QString>
#include <QStringList>
#include <QVariant>
#include <QAxBase>
#include <QAxObject>

#include <QTableWidget>
#include <QAbstractItemModel>
#include <QDebug>

#include "tablemodel.h"

typedef unsigned int quint32;

/**
  *@brief 这是一个便于Qt读写excel封装的类，同时，便于把excel中的数据
  *显示到界面上，或者把界面上的数据写入excel中，同GUI进行交互，关系如下：
  * DB --〉QSqlTableModel --> CExcelEngine <--> xlsx file.
  *						         |
  * QTableWidget <--------------/
  *@note CExcelEngine类只负责读/写数据，不负责解析，做中间层
  */
class CExcelEngine : protected QObject
{
public:
    CExcelEngine();
    CExcelEngine(const QString& xlsFile);
    ~CExcelEngine();

public:
	//打开xlsx文件。要提前设置excel文件
    bool open(quint32 nSheet = 1, bool visible = false);
    bool open(const QString& xlsFile, quint32 nSheet = 1, bool visible = false);
    void save();                //保存xlsx报表
    void close();               //关闭xlsx报表

    bool saveDataFromTable(QAbstractItemModel *model);	//保存数据到xlsx
    bool readDataToTable(QTableWidget *tableWidget);	//从xls读取数据到ui
	bool readDataToTable(CTableModel *model);

    QVariant getCellData(quint32 row, quint32 column);                //获取指定单元数据
    bool     setCellData(quint32 row, quint32 column, QVariant data); //修改指定单元数据

	void setColumnWidth(int column, int width);
	void setRowHeight(int row, int height);

    quint32 rowCount()const;
    quint32 columnCount()const;

	void setCellTextCenter(int row, int column);
	void setCellTextCenter(const QString &cell);
	void setAllCellTextCenter();

	void setCellTextWrap(int row, int column, bool isWrap);
	void setCellTextWrap(const QString &cell, bool isWrap);

	void setAutoFitColumn(int from, int to);
	void setAllColumnAutoFit();
	
	void setCellFontBold(const QString &cell, bool isBold);

	void getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn);

	void clearCell(int row, int column);
	void clearRangeCells(int rowFrom, int columnFrom, int rowTo, int columnTo);

    bool isOpen();
    bool isValid();

protected:
    void clear();

private:
	void init();

private:
    QAxObject *m_pExcel;      //指向整个excel应用程序
    QAxObject *m_pWorkbooks;  //指向工作簿集,excel有很多工作簿
    QAxObject *m_pWorkbook;   //指向sXlsFile对应的工作簿
    QAxObject *m_pWorksheet;  //指向工作簿中的某个sheet表单

    QString   m_sXlsFile;     //xls文件路径
    int       m_nRowCount;    //行数
    int       m_nColumnCount; //列数
    int       m_nStartRow;    //开始有数据的行下标值
    int       m_nStartColumn; //开始有数据的列下标值

    bool      m_bIsOpen;      //是否已打开
    bool      m_bIsValid;     //是否有效
    bool      m_bNewFile;  //是否是一个新建xls文件，用来区分打开的excel是已存在文件还是有本类新建的
};

#endif // EXCELENGINE_H
