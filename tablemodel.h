#ifndef SERVERLISTMODEL_H
#define SERVERLISTMODEL_H

#include <QAbstractTableModel>

class CTableModel : public QAbstractTableModel
{
	Q_OBJECT

public:
	explicit CTableModel(QObject *parent = NULL);
	~CTableModel();

	int rowCount(const QModelIndex &parent  = QModelIndex()) const Q_DECL_OVERRIDE;
	int	columnCount(const QModelIndex & parent = QModelIndex()) const Q_DECL_OVERRIDE;

	QVariant data(const QModelIndex &index, int role  = Qt::DisplayRole) const Q_DECL_OVERRIDE;
	bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) Q_DECL_OVERRIDE;

	QVariant headerData(int section, Qt::Orientation orientation,
		int role = Qt::DisplayRole) const Q_DECL_OVERRIDE;
	bool setHeaderData(int section, Qt::Orientation orientation, const QVariant &value,
		int role = Qt::EditRole) Q_DECL_OVERRIDE;

	Qt::ItemFlags flags(const QModelIndex &index) const Q_DECL_OVERRIDE;

	bool insertRows(int row, int count, const QModelIndex & parent = QModelIndex()) Q_DECL_OVERRIDE;
	void setHorizontalHeaders(const QStringList&);

	void reset();

private:
	QString getKeyString(const QModelIndex &index) const;

private:
	int m_rowCount;

	QMap<QString, QString> m_mapData;
	QMap<int, QString> m_mapHeader;
};

#endif //GAMELISTMODEL_H