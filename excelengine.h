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
  *@brief ����һ������Qt��дexcel��װ���࣬ͬʱ�����ڰ�excel�е�����
  *��ʾ�������ϣ����߰ѽ����ϵ�����д��excel�У�ͬGUI���н�������ϵ���£�
  * DB --��QSqlTableModel --> CExcelEngine <--> xlsx file.
  *						         |
  * QTableWidget <--------------/
  *@note CExcelEngine��ֻ�����/д���ݣ���������������м��
  */
class CExcelEngine : protected QObject
{
public:
    CExcelEngine();
    CExcelEngine(const QString& xlsFile);
    ~CExcelEngine();

public:
	//��xlsx�ļ���Ҫ��ǰ����excel�ļ�
    bool open(quint32 nSheet = 1, bool visible = false);
    bool open(const QString& xlsFile, quint32 nSheet = 1, bool visible = false);
    void save();                //����xlsx����
    void close();               //�ر�xlsx����

    bool saveDataFromTable(QAbstractItemModel *model);	//�������ݵ�xlsx
    bool readDataToTable(QTableWidget *tableWidget);	//��xls��ȡ���ݵ�ui
	bool readDataToTable(CTableModel *model);

    QVariant getCellData(quint32 row, quint32 column);                //��ȡָ����Ԫ����
    bool     setCellData(quint32 row, quint32 column, QVariant data); //�޸�ָ����Ԫ����

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
    QAxObject *m_pExcel;      //ָ������excelӦ�ó���
    QAxObject *m_pWorkbooks;  //ָ��������,excel�кܶ๤����
    QAxObject *m_pWorkbook;   //ָ��sXlsFile��Ӧ�Ĺ�����
    QAxObject *m_pWorksheet;  //ָ�������е�ĳ��sheet��

    QString   m_sXlsFile;     //xls�ļ�·��
    int       m_nRowCount;    //����
    int       m_nColumnCount; //����
    int       m_nStartRow;    //��ʼ�����ݵ����±�ֵ
    int       m_nStartColumn; //��ʼ�����ݵ����±�ֵ

    bool      m_bIsOpen;      //�Ƿ��Ѵ�
    bool      m_bIsValid;     //�Ƿ���Ч
    bool      m_bNewFile;  //�Ƿ���һ���½�xls�ļ����������ִ򿪵�excel���Ѵ����ļ������б����½���
};

#endif // EXCELENGINE_H
