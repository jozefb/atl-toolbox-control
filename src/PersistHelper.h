// PersistHelper.h

#pragma once

template<class T>
class CPersistHelper
{

public:

	CPersistHelper() : m_bRequiresSave(false)
	{

	}

	void SetDirty(bool fDirty)
	{
		m_bRequiresSave=true;
	}
	unsigned m_bRequiresSave:1;
};
