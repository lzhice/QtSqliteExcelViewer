#ifndef PWDDLG_H
#define PWDDLG_H

#include <QDialog>
#include <QLineEdit>
#include <QPushButton>
#include <QLabel>
#include <QPoint>
#include <QMouseEvent>

/*
//  Function: 验证密码消息框
*/
class  CPwdDlg : public QDialog
{
	Q_OBJECT

public:
	CPwdDlg(QWidget *parent = NULL);
	~CPwdDlg();

	QString getPassword() const;

protected:
	virtual void mousePressEvent(QMouseEvent *event);
	virtual void mouseReleaseEvent(QMouseEvent *event);
	virtual void mouseMoveEvent(QMouseEvent *event);
	virtual void paintEvent(QPaintEvent *event);

private:
	void initCtrls();

private:
	QPushButton *m_pBtnEnsure;
	QLineEdit *m_pLineEdit;

	QPoint m_ptLastPos;
	bool m_bIsLeftPress;
};

#endif // PWDDLG_H
