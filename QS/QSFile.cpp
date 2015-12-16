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

STDMETHODIMP CQSFile::Open(BSTR bstrPath, VARIANT_BOOL* pbOk)
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
		GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE,
		NULL,
		OPEN_EXISTING,
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
