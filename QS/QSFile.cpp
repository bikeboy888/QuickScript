// QSFile.cpp : Implementation of CQSFile

#include "stdafx.h"
#include "QSFile.h"

// CQSFile

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSFile
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Clone(IStream **ppstm)
{
	return E_NOTIMPL;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Close()
{
	HRESULT hr = S_OK;
	BOOL bOk = TRUE;
	if (m_hFile == INVALID_HANDLE_VALUE)
	{
		return S_OK;
	}

	bOk = CloseHandle(m_hFile);
	if (bOk != TRUE)
	{
		return S_FALSE;
	}

	m_hFile = INVALID_HANDLE_VALUE;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Commit(DWORD grfCommitFlags)
{
	return E_NOTIMPL;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::CopyTo(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten)
{
	return E_NOTIMPL;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Open(BSTR bstrPath, LONG nDesiredAccess, LONG nShareMode, LONG nCreationDisposition, VARIANT_BOOL* pbOk)
{
	HRESULT hr = S_OK;

	if (pbOk == NULL) return E_INVALIDARG;
	*pbOk = VARIANT_FALSE;

	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		CHECKHR(Close());
	}

	m_hFile = CreateFile(
		bstrPath,
		(DWORD) nDesiredAccess,
		(DWORD) nShareMode,
		NULL,
		(DWORD) nCreationDisposition,
		FILE_ATTRIBUTE_NORMAL,
		NULL);

	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		*pbOk = VARIANT_TRUE;
		return S_OK;
	}

	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Read(void *pv, ULONG cb, ULONG *pcbRead)
{
	HRESULT hr = S_OK;
	BOOL bOk = TRUE;
	DWORD dwRead = 0;
	if (pcbRead) *pcbRead = 0;
	if (m_hFile == INVALID_HANDLE_VALUE) return S_FALSE;
	bOk = ReadFile(m_hFile, pv, cb, &dwRead, NULL);
	if (bOk != TRUE)
	{
		return S_FALSE;
	}
	if (pcbRead) *pcbRead = dwRead;
	if (dwRead < cb)
	{
		return S_FALSE;
	}
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::ReadAll(BSTR* pbstrText)
{
	HRESULT hr = S_OK;

	if (!pbstrText) return E_INVALIDARG;
	*pbstrText = NULL;

	if (m_hFile == INVALID_HANDLE_VALUE) return S_FALSE;

	DWORD dwSize = GetFileSize(m_hFile, NULL);
	if (dwSize == 0) return S_OK;

	DWORD dwPos = SetFilePointer(m_hFile, 0, NULL, FILE_BEGIN);
	int nLenA = dwSize + 1;
	CComHeapPtr<BYTE> pDataA;
	pDataA.Allocate(nLenA);
	ZeroMemory((BYTE*) pDataA, nLenA);
	DWORD dwRead = 0;
	BOOL bOk = ReadFile(m_hFile, (BYTE*) pDataA, nLenA, &dwRead, NULL);
	if (bOk != TRUE)
	{
		return S_FALSE;
	}

	UINT CodePage = m_CodePage;
	int nLenW = MultiByteToWideChar(CodePage, 0, (LPCSTR) (BYTE*) pDataA, nLenA, NULL, 0);
	if (nLenW <= 1)
	{
		return S_FALSE;
	}
	BSTR pDataW = SysAllocStringLen(NULL, nLenW);
	MultiByteToWideChar(CodePage, 0, (LPCSTR) (BYTE*) pDataA, nLenA, pDataW, nLenW);
	*pbstrText = pDataW;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition)
{
	HRESULT hr = S_OK;
	DWORD dwLastError = 0;
	DWORD dwMethod = 0;
	DWORD dwPos = 0;
	if (m_hFile == INVALID_HANDLE_VALUE) return S_FALSE;
	switch (dwOrigin)
	{
	case STREAM_SEEK_CUR:
		dwMethod = FILE_CURRENT;
		break;
	case STREAM_SEEK_END:
		dwMethod = FILE_END;
		break;
	case STREAM_SEEK_SET:
		dwMethod = FILE_BEGIN;
		break;
	default:
		return E_INVALIDARG;
	}
	dwPos = SetFilePointer(m_hFile, dlibMove.LowPart, &dlibMove.HighPart, dwMethod);
	if (dwPos == INVALID_SET_FILE_POINTER)
	{
		dwLastError = GetLastError();
		if (dwLastError != ERROR_SUCCESS)
		{
			return S_FALSE;
		}
	}
	plibNewPosition->LowPart = dwPos;
	plibNewPosition->HighPart = dlibMove.HighPart;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::SetSize(ULARGE_INTEGER libNewSize)
{
	return E_NOTIMPL;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Stat(STATSTG *pstatstg, DWORD grfStatFlag)
{
	HRESULT hr = S_OK;
	DWORD dwLastError = ERROR_SUCCESS;
	if (!pstatstg) return E_INVALIDARG;
	ZeroMemory(pstatstg, sizeof(STATSTG));
	pstatstg->cbSize.LowPart = GetFileSize(m_hFile, &pstatstg->cbSize.HighPart);
	if (pstatstg->cbSize.LowPart == INVALID_FILE_SIZE)
	{
		dwLastError = GetLastError();
		if (dwLastError != NO_ERROR)
		{
			return S_FALSE;
		}
	}
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSFile::Write(const void *pv, ULONG cb, ULONG *pcbWritten)
{
	HRESULT hr = S_OK;
	BOOL bOk = TRUE;
	DWORD dwWritten = 0;
	if (pcbWritten) *pcbWritten = 0;
	if (m_hFile == INVALID_HANDLE_VALUE) return S_FALSE;
	bOk = WriteFile(m_hFile, pv, cb, &dwWritten, NULL);
	if (bOk != TRUE)
	{
		return S_FALSE;
	}
	if (pcbWritten) *pcbWritten = dwWritten;
	if (dwWritten < cb)
	{
		return S_FALSE;
	}
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
