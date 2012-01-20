// DataObjectEx.cpp: implementation of the CDataObjectEx class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DataObjectEx.h"
#include "EnumFormatEtc.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////


HRESULT _CopyStgMedium(STGMEDIUM* pMedDest, STGMEDIUM* pMedSrc, FORMATETC* pFmtSrc)
{
	switch( pMedSrc->tymed ) {
	case TYMED_GDI:
	case TYMED_FILE:
	case TYMED_ENHMF:
	case TYMED_MFPICT:
	case TYMED_HGLOBAL:
		pMedDest->hGlobal = (HGLOBAL) ::OleDuplicateData(pMedSrc->hGlobal, pFmtSrc->cfFormat, NULL);
		break;
	case TYMED_ISTREAM:
		{
            pMedDest->pstm = NULL;
            if( FAILED( ::CreateStreamOnHGlobal(NULL, TRUE, &pMedDest->pstm) ) ) return E_OUTOFMEMORY;
			LARGE_INTEGER dlibMove = { 0, 0 };
			ATLVERIFY(SUCCEEDED(pMedSrc->pstm->Seek(dlibMove, STREAM_SEEK_SET, NULL)));
            ULARGE_INTEGER alot = { 0, INT_MAX };
            if( FAILED( pMedSrc->pstm->CopyTo(pMedDest->pstm, alot, NULL, NULL) ) )
			{
				ATLASSERT(FALSE);
				return E_FAIL;
			}
		}
		break;
	case TYMED_ISTORAGE:
		{
            pMedDest->pstg = NULL;
            if( FAILED( ::StgCreateDocfile(NULL, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DELETEONRELEASE | STGM_CREATE, 0, &pMedDest->pstg) ) ) return E_OUTOFMEMORY;
            if( FAILED( pMedSrc->pstg->CopyTo(0, NULL, NULL, pMedDest->pstg) ) ) return E_FAIL;
		}
		break;
	default:
		ATLASSERT(false);
		return DV_E_TYMED;
	}
	pMedDest->tymed = pMedSrc->tymed;
	pMedDest->pUnkForRelease = pMedSrc->pUnkForRelease;
	return S_OK;
}



