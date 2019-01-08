#include "tablemodel.h"
#include <QSize>

CTableModel::CTableModel(QObject *parent)
	: QAbstractTableModel(parent),
	m_rowCount(0)
{
}

CTableModel::~CTableModel()
{
}

bool CTableModel::insertRows(int row, int count, const QModelIndex & parent/* = QModelIndex()*/)
{
	if (row >= 0 && count > 0)
	{
		beginInsertRows(QModelIndex(), row, row+count-1);
		m_rowCount = rowCount() + count;
		endInsertRows();
		return true;
	}
	return false;
}

void CTableModel::setHorizontalHeaders(const QStringList& strlstHeaders)
{
	m_mapHeader.clear();
	beginInsertColumns(QModelIndex(), 0, strlstHeaders.size()-1);
	for (int i = 0; i < strlstHeaders.size(); ++i)
	{
		m_mapHeader.insert(i, strlstHeaders[i]);
	}
	endInsertColumns();
}

int CTableModel::rowCount(const QModelIndex &parent  /*= QModelIndex()*/) const
{
	return m_rowCount;
}

int CTableModel::columnCount(const QModelIndex &parent  /*= QModelIndex()*/) const
{
	return m_mapHeader.size();
}

QString CTableModel::getKeyString(const QModelIndex &index) const
{
	return QString("%1:%2").arg(index.row()).arg(index.column());
}

QVariant CTableModel::data(const QModelIndex &index, int role  /*= Qt::DisplayRole*/) const
{
	if ( index.isValid() && role == Qt::DisplayRole)
	{
		QString key = getKeyString(index);
		if (m_mapData.contains(key))
		{
			return QVariant::fromValue(m_mapData[key]);
		}
	}
	return QVariant();
}

bool CTableModel::setData(const QModelIndex &index, const QVariant &value, int role/* = Qt::EditRole*/)
{
	if (!index.isValid() || role != Qt::EditRole)
	{
		return false;
	}

	if (role != Qt::EditRole)
	{
		return QAbstractTableModel::setData(index, value, role);
	}

	m_mapData.insert(getKeyString(index), value.toString());
	return true;
}

QVariant CTableModel::headerData(int section, Qt::Orientation orientation, int role/* = Qt::DisplayRole*/) const
{
	if (role != Qt::DisplayRole)
	{
		return QVariant();
	}

	if (orientation == Qt::Horizontal)
	{
		if (m_mapHeader.contains(section))
		{
			return QVariant::fromValue(m_mapHeader.value(section));
		}
	}
	else
	{
		return QVariant(section+1);
	}
	return QVariant();
}

bool CTableModel::setHeaderData(int section, Qt::Orientation orientation, const QVariant &value, int role/* = Qt::EditRole*/)
{
	if (role != Qt::EditRole)
	{
		return false;
	}

	m_mapHeader.insert(section, value.toString());
	return true;
}

Qt::ItemFlags CTableModel::flags(const QModelIndex &index) const
{
	if (!index.isValid())
	{
		return 0;
	}
	return QAbstractTableModel::flags(index) | Qt::ItemIsEditable;
}

void CTableModel::reset()
{
	beginResetModel();
	m_mapData.clear();
	m_rowCount = 0;
	endResetModel();
}