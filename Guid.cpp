// Guid.cpp: implementation of the CGuid class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Guid.h"
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

#pragma comment(lib, "rpcrt4.lib")

#define GUID_STRING_LEN	40
 
CGuid::CGuid()
{
	m_guid = GUID_NULL;
}

CGuid::CGuid(BSTR guid)
{
	this->operator=(CComBSTR(guid));
}

CGuid::~CGuid()
{

}

CGuid::CGuid(const CGuid& g)
{
	m_guid = g.m_guid;	
}

CGuid::CGuid(const GUID& g)
{
	m_guid = g;
}

bool CGuid::operator==(const GUID& g)
{
	return !::memcmp(&m_guid, &g, sizeof(GUID));
}

bool CGuid::operator==(const CGuid& g)
{	
	return !::memcmp(&m_guid, &g.m_guid, sizeof(GUID));
}

CGuid::operator GUID&()
{ 
	return m_guid; 
}

CGuid::operator GUID*()
{ 
	return &m_guid; 
}

CGuid::operator CComBSTR()
{ 
	CComBSTR bstrGuid;
	Convert(m_guid, bstrGuid);
	return bstrGuid; 
}

CGuid::operator CString()
{ 
	CString csGuid;
	Convert(m_guid, csGuid);
	return csGuid; 
}

bool CGuid::operator!=(const GUID& g)
{
	return ::memcmp(&m_guid, &g, sizeof(GUID)) != 0;
}

bool CGuid::operator!=(const CGuid& g)
{
	return ::memcmp(&m_guid, &g.m_guid, sizeof(GUID)) != 0;
}

CGuid& CGuid::operator=(const GUID& g)
{
	if( ::memcmp(&m_guid, &g, sizeof(GUID)) != 0 )
	{
		copy(g);
	}

	return *this;
}

CGuid& CGuid::operator=(const CComBSTR& g)
{
	Convert(g, m_guid);
	return *this;
}

CGuid& CGuid::operator=(const CString& g)
{
	Convert(g, m_guid);
	return *this;
}

CGuid& CGuid::operator=(const CGuid& g)
{
	if(this != &g )
		copy(g.m_guid);
	return *this;
}

inline void CGuid::copy(const CGuid& g)
{
	::memcpy(&m_guid, (void*)&g.m_guid, sizeof(GUID));
}

inline void CGuid::copy(const GUID& g)
{
	::memcpy(&m_guid, (void*)&g, sizeof(GUID));
}

bool CGuid::operator<(const CGuid& g1) const
{
	RPC_STATUS status;
	return ::UuidCompare(const_cast<GUID*>(&m_guid), const_cast<GUID*>(&g1.m_guid), &status)==-1;
}

bool CGuid::Convert(GUID& guid, CComBSTR& bstrGuid)
{
	OLECHAR szGuid[GUID_STRING_LEN]={0};
	::StringFromGUID2(guid, szGuid, GUID_STRING_LEN);
	bstrGuid = szGuid;
	
	return true;
}

bool CGuid::Convert(GUID& guid, CString& csGuid)
{
	OLECHAR szGuid[GUID_STRING_LEN]={0};
	HRESULT hr = ::StringFromGUID2(guid, szGuid, GUID_STRING_LEN);
	csGuid = szGuid;
	
	return true;
}

bool CGuid::Convert(const CComBSTR& bstrGuid, GUID& guid)
{
	::ZeroMemory(&guid, sizeof(GUID));
	CString csguid=bstrGuid;	
	return ::UuidFromString(reinterpret_cast<unsigned char*>(csguid.GetBuffer(GUID_STRING_LEN)), &guid) == RPC_S_OK;
}

bool CGuid::Convert(const CString& csGuid, GUID& guid)
{
	::ZeroMemory(&guid, sizeof(GUID));	
	CString* pG = const_cast<CString*>(&csGuid);
	return ::UuidFromString(reinterpret_cast<unsigned char*>(pG->GetBuffer(GUID_STRING_LEN)), &guid) == RPC_S_OK;
}


bool CGuid::Create(CComBSTR& bstrGuid)
{
	GUID guid;
	if( !Create(guid) )
	{
		return false;
	}

	return Convert(guid, bstrGuid);
}

bool CGuid::Create(GUID& guid)
{
	return ::UuidCreate(&guid) == RPC_S_OK;
}

long CGuid::HashKey(GUID& guid)
{
	RPC_STATUS status;
	return ::UuidHash(&guid, &status);
}

BOOL CGuid::ProgID(CString& csProgID)
{
	CComBSTR bstrTmp;
	BOOL bRet = ProgID(bstrTmp);
	csProgID = bstrTmp;
	return bRet;
}
	
BOOL CGuid::ProgID(CComBSTR& bstrProgID)
{
	BOOL bRet=FALSE;
	if( *this == GUID_NULL )
	{
		return bRet;
	}

	LPOLESTR psz = NULL;
	if( ::ProgIDFromCLSID(m_guid, &psz) == S_OK )
	{
		bstrProgID=psz;
		::CoTaskMemFree(psz);
		psz=NULL;
		bRet=TRUE;
	}
	return bRet;
}
