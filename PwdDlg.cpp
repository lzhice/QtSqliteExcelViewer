#include <QPainter>
#include "pwddlg.h"

CPwdDlg::CPwdDlg(QWidget *parent)
	: QDialog(parent),
	m_bIsLeftPress(false),
	m_ptLastPos(QPoint(0, 0))
{
	setWindowFlags(windowFlags() | Qt::FramelessWindowHint);

	resize(300, 100);
	setStyleSheet("QDialog{ \
				  background-color: rgb(215, 215, 215); \
				  color: rgb(123, 123, 123); \
				  font-size: 12px; }");
	initCtrls();
}

CPwdDlg::~CPwdDlg()
{
}

void CPwdDlg::initCtrls()
{
	//文本框
	m_pLineEdit = new QLineEdit(this);
	m_pLineEdit->setText("");
	m_pLineEdit->setEchoMode(QLineEdit::Password);
	m_pLineEdit->setGeometry(20, 61, 160, 23);
	m_pLineEdit->setStyleSheet("QLineEdit{color: rgb(123, 123, 123);font-size: 13px;}");
	m_pLineEdit->setText("7F_QX_DB_ACCOUNT");

	m_pBtnEnsure = new QPushButton(this);
	m_pBtnEnsure->setText(QString::fromStdWString(L"确定"));
	m_pBtnEnsure->setGeometry(200, 61, 76, 23);

	connect(m_pBtnEnsure, SIGNAL(clicked()), this, SLOT(accept()));
}


void CPwdDlg::mousePressEvent(QMouseEvent *event)
{
	if(event->button() == Qt::LeftButton)
	{
		m_bIsLeftPress = true;
		m_ptLastPos = event->globalPos();
	}
}

void CPwdDlg::mouseMoveEvent(QMouseEvent *event)
{
	if (m_bIsLeftPress)
	{
		QPoint ptOffset = event->globalPos() - m_ptLastPos;
		m_ptLastPos = event->globalPos();
		move(ptOffset + pos());	
	}
}

void CPwdDlg::mouseReleaseEvent(QMouseEvent *event)
{
	if (event->button() == Qt::LeftButton)
	{
		m_bIsLeftPress = false;
	}
}

void CPwdDlg::paintEvent(QPaintEvent *event)
{
	QPainter painter(this);
	painter.drawText(20, 40, QString::fromStdWString(L"请输入密码："));
}

QString CPwdDlg::getPassword() const
{
	return m_pLineEdit->text();
}