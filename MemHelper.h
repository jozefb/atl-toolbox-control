// MemHelper.h

#pragma once

#define _countof(array) (sizeof(array)/sizeof(array[0]))

#define INITPTR(pObj)	(pObj = NULL)
#define DELETEARRAY(pObj) if ( pObj     != NULL ) { delete [] pObj;        pObj     = NULL; }
#define DELETEP(pObj)    if ( pObj     != NULL ) { delete pObj;        pObj     = NULL; }
#define FREE(pHeapMem)   if ( pHeapMem != NULL ) { free(pHeapMem);     pHeapMem = NULL; }
#define RELEASE(pIfc) if ( pIfc != NULL ) { pIfc->Release(); pIfc=NULL; }
#define ADDREF(pIfc) if ( pIfc != NULL ) { pIfc->AddRef(); }
#define FREEBSTR(bstr)   if ( bstr	 != NULL ) { SysFreeString(bstr);bstr	 = NULL; }
#define SETFORMATETC(fe, cf, asp, td, med, li)   \
    ((fe).cfFormat=cf, \
    (fe).dwAspect=asp, \
    (fe).ptd=td, \
    (fe).tymed=med, \
	(fe).lindex=li)

#define COMEMALLOC(cnt) ::CoTaskMemAlloc(cnt)
#define COMEMREALLOC(pData, cnt) ::CoTaskMemRealloc(pData, cnt);
#define COMEMFREE(pData) if(pData != NULL) {::CoTaskMemFree(pData); pData = NULL; }

#define FREERES(hRes) { if(hRes) { ::FreeResource(hRes); hRes = NULL;}}


template<class T>
class CoMemItemPtr
{
public:
	CoMemItemPtr() : p(NULL){
	}

	~CoMemItemPtr(){
		COMEMFREE(p);
	}

	T* p;	

	T& operator&(){
		return &p;
	}

	T operator[](int index){
		return p[index];
	}
};

