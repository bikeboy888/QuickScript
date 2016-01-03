// QSNet.cpp : Implementation of CQSNet

#include "stdafx.h"
#include "QSNet.h"
#include "QSFile.h"
#include "..\QSUtil\OutputDebugFormat.h"
#include "..\QSUtil\OutputDebugRequest.h"
#include "..\QSUtil\VariantToBOOL.h"
#include "..\QSUtil\InvokeMethod.h"

// CQSNet

STDMETHODIMP CQSNet::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSNet
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

CQSNet::CQSNet() :
	m_bAsync(VARIANT_FALSE),
	m_hInternet(NULL),
	m_hConnect(NULL),
	m_hRequest(NULL),
	m_hOpenEvent(NULL),
	m_nOpenTimeout(10000),
	m_state(stateUnknown),
	m_nReadyState(0)
{
	ZeroMemory(&m_URLComponents, sizeof(m_URLComponents));
	ZeroMemory(&m_szScheme, sizeof(m_szScheme));
	ZeroMemory(&m_szHostName, sizeof(m_szHostName));
	ZeroMemory(&m_szUserName, sizeof(m_szUserName));
	ZeroMemory(&m_szPassword, sizeof(m_szPassword));
	ZeroMemory(&m_szUrlPath, sizeof(m_szUrlPath));
	ZeroMemory(&m_szExtraInfo, sizeof(m_szExtraInfo));
	m_ResponseBody.Create((ULONG) 0);
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_OnReadyStateChange(VARIANT* pvarDispatch)
{
	HRESULT hr = S_OK;
	VariantInit(pvarDispatch);
	if (!m_OnReadyStateChange) return S_OK;
	CComPtr<IDispatch> spIDispatch = m_OnReadyStateChange;
	pvarDispatch->vt = VT_DISPATCH;
	V_DISPATCH(pvarDispatch) = spIDispatch.Detach();
	/*
	if (!ppIDispatch) return E_INVALIDARG;
	*ppIDispatch = NULL;
	if (!m_OnReadyStateChange) return S_OK;
	CHECKHR(m_OnReadyStateChange->QueryInterface(IID_IDispatch, (void**) ppIDispatch));
	*/
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::put_OnReadyStateChange(VARIANT pvarDispatch)
{
	//m_OnReadyStateChange = pIDispatch;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::putref_OnReadyStateChange(VARIANT varDispatch)
{
	HRESULT hr = S_OK;
	VARIANT* pvarDispatch = &varDispatch;
	switch (pvarDispatch->vt)
	{
	case VT_DISPATCH:
		m_OnReadyStateChange = NULL;
		CHECKHR(V_DISPATCH(pvarDispatch)->QueryInterface(IID_IDispatch, (void**) &m_OnReadyStateChange));
		return hr;
	default:
		return E_NOTIMPL;
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_OpenTimeout(LONG* pnOpenTimeout)
{
	if (!pnOpenTimeout) return E_INVALIDARG;
	*pnOpenTimeout = m_nOpenTimeout;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::put_OpenTimeout(LONG nOpenTimeout)
{
	m_nOpenTimeout = nOpenTimeout;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_ReadyState(LONG* pnReadyState)
{
	if (!pnReadyState) return E_INVALIDARG;
	*pnReadyState = m_nReadyState;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::put_ReadyState(LONG nReadyState)
{
	HRESULT hr = S_OK;

	if (m_nReadyState == nReadyState)
	{
		return S_OK;
	}

	m_nReadyState = nReadyState;
#ifdef _DEBUG
	OutputDebugFormat(L"ReadyState %d\r\n", nReadyState);
#endif
	CHECKHR(OnReadyStateChange());
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_ResponsePath(BSTR* pbstrResponsePath)
{
	if (!pbstrResponsePath) return E_INVALIDARG;
	*pbstrResponsePath = m_bstrResponsePath.Copy();
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::put_ResponsePath(BSTR bstrResponsePath)
{
	m_bstrResponsePath = bstrResponsePath;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_ResponseText(BSTR* pbstrText)
{
	HRESULT hr = S_OK;
	if (!pbstrText) return E_INVALIDARG;
	*pbstrText = NULL;

	int nLenA = m_ResponseBody.GetCount() + 1;
	CHECKHR(m_ResponseBody.Resize(nLenA + 1));

	int nLenW = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR) &m_ResponseBody[0], nLenA, NULL, NULL);
	*pbstrText = SysAllocStringLen(NULL, nLenW - 1);
	if (*pbstrText == NULL)
	{
		m_ResponseBody.Resize(nLenA - 1);
		return E_OUTOFMEMORY;
	}

	MultiByteToWideChar(CP_UTF8, 0, (LPCSTR) &m_ResponseBody[0], m_ResponseBody.GetCount(), *pbstrText, nLenW);
	CHECKHR(m_ResponseBody.Resize(nLenA - 1));

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::get_Status(LONG* pnStatus)
{
	HRESULT hr = S_OK;
	BOOL bOk = FALSE;
	DWORD dwLastError = ERROR_SUCCESS;
	if (!pnStatus) return E_INVALIDARG;
	*pnStatus = 0;
	if (!m_hRequest) return S_FALSE;

	WORD statusCode = 0;
	DWORD length = sizeof(DWORD);
	bOk = HttpQueryInfo(
		m_hRequest,
		HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER,
		&statusCode,
		&length,
		NULL);
	if (bOk != TRUE)
	{
		dwLastError = GetLastError();
		return S_OK;
	}
	*pnStatus = statusCode;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::Close()
{
	HRESULT hr = S_OK;
	BOOL bOk = TRUE;
	DWORD dwLastError = ERROR_SUCCESS;

	if (m_bAsync == VARIANT_TRUE && m_hInternet != NULL)
	{
		InternetSetStatusCallback(m_hInternet, NULL);
	}

	if (m_hOpenEvent)
	{
		bOk = CloseHandle(m_hOpenEvent);
		if (bOk != TRUE)
		{
#ifdef _DEBUG
			dwLastError = GetLastError();
			OutputDebugFormat(L"CloseHandle(m_hOpenEvent) = FALSE (GetLastError:%d)\r\n", dwLastError);
#endif
			return S_FALSE;
		}
		m_hOpenEvent = NULL;
	}

	if (m_hRequest)
	{
		bOk = InternetCloseHandle(m_hRequest);
		if (bOk != TRUE)
		{
#ifdef _DEBUG
			dwLastError = GetLastError();
			OutputDebugFormat(L"InternetCloseHandle(m_hRequest) = FALSE (GetLastError:%d)\r\n", dwLastError);
#endif
			return S_FALSE;
		}
		m_hRequest = NULL;
	}
	if (m_hConnect)
	{
		bOk = InternetCloseHandle(m_hConnect);
		if (bOk != TRUE)
		{
#ifdef _DEBUG
			dwLastError = GetLastError();
			OutputDebugFormat(L"InternetCloseHandle(m_hConnect) = FALSE (GetLastError:%d)\r\n", dwLastError);
#endif
			return S_FALSE;
		}
		m_hConnect = NULL;
	}

	if (m_hInternet)
	{
		bOk = InternetCloseHandle(m_hInternet);
		if (bOk != TRUE)
		{
#ifdef _DEBUG
			dwLastError = GetLastError();
			OutputDebugFormat(L"InternetCloseHandle(m_hInternet) = FALSE (GetLastError:%d)\r\n", dwLastError);
#endif
			return S_FALSE;
		}
		m_hInternet = NULL;
	}

#ifdef X
	DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE;
	if (m_URLComponents.nScheme == INTERNET_SCHEME_HTTPS)
	{
#ifdef UNDER_CE
		dwFlags |= INTERNET_FLAG_IGNORE_CERT_CN_INVALID;
		dwFlags |= INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;
#endif
		dwFlags |= INTERNET_FLAG_SECURE;
	}
	dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;

	//m_asyncState = asyncStateHttpOpenRequest;
	m_hRequest = HttpOpenRequest(
		m_hConnect,
		m_bstrMethod,
		bstrObjectName,
		NULL,
		NULL,
		NULL,
		dwFlags,
		(DWORD_PTR) this);

	if (m_hRequest == NULL)
	{
		dwLastError = GetLastError();
		if (dwLastError != ERROR_IO_PENDING)
		{
			hr = HRESULT_FROM_WIN32(dwLastError);
			return hr;
		}
		if (m_bAsync == VARIANT_TRUE)
		{
			WaitForSingleObject(m_hRequestOpenedEvent, INFINITE);
		}
	}
#endif

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoInternetOpen()
{
	HRESULT hr = S_OK;
	DWORD dwLastError = ERROR_SUCCESS;

	CComBSTR bstrUserAgent(L"QuickScript");

	m_state = stateInternetOpen;

	m_hInternet = InternetOpen(
		(BSTR) bstrUserAgent,
		INTERNET_OPEN_TYPE_PRECONFIG,
		NULL,
		NULL,
		(m_bAsync == VARIANT_TRUE) ? INTERNET_FLAG_ASYNC : 0 );
#ifdef _DEBUG
	OutputDebugFormat(L"m_hInternet = 0x%p\r\n", m_hInternet);
#endif
	if (m_hInternet == NULL)
	{
		dwLastError = GetLastError();
		hr = HRESULT_FROM_WIN32(dwLastError);
		return hr;
	}

	if (m_bAsync == VARIANT_TRUE)
	{
		if (InternetSetStatusCallback(m_hInternet, InternetStatusCallback) == INTERNET_INVALID_STATUS_CALLBACK)
		{
			dwLastError = GetLastError();
			hr = HRESULT_FROM_WIN32(dwLastError);
			return hr;
		}
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoInternetConnect()
{
	HRESULT hr = S_OK;
	DWORD dwLastError = 0;

	CComBSTR bstrUserAgent(L"QuickScript");

	ZeroMemory(&m_URLComponents, sizeof(m_URLComponents));
	m_URLComponents.dwStructSize = sizeof(URL_COMPONENTS);
	m_URLComponents.lpszScheme = m_szScheme;
	m_URLComponents.dwSchemeLength = sizeof(m_szScheme) / sizeof(TCHAR) - 1;
	m_URLComponents.lpszHostName = m_szHostName;
	m_URLComponents.dwHostNameLength = sizeof(m_szHostName) / sizeof(TCHAR) - 1;
	m_URLComponents.lpszUserName = m_szUserName;
	m_URLComponents.dwUserNameLength = sizeof(m_szUserName) / sizeof(TCHAR) - 1;
	m_URLComponents.lpszPassword = m_szPassword;
	m_URLComponents.dwPasswordLength = sizeof(m_szPassword) / sizeof(TCHAR) - 1;
	m_URLComponents.lpszUrlPath = m_szUrlPath;
	m_URLComponents.dwUrlPathLength = sizeof(m_szUrlPath) / sizeof(TCHAR) - 1;
	m_URLComponents.lpszExtraInfo = m_szExtraInfo;
	m_URLComponents.dwExtraInfoLength = sizeof(m_szExtraInfo) / sizeof(TCHAR) - 1;

	::InternetCrackUrl(m_bstrURL, m_bstrURL.Length(), ICU_DECODE, &m_URLComponents);

	DWORD dwService = INTERNET_SERVICE_HTTP;
	switch (m_URLComponents.nScheme)
	{
	case INTERNET_SCHEME_HTTP:
		dwService = INTERNET_SERVICE_HTTP;
		break;

	case INTERNET_SCHEME_HTTPS:
		dwService = INTERNET_SERVICE_HTTP;
		break;
	}

	m_state = stateInternetConnect;

	m_hConnect = InternetConnect(
		m_hInternet,
		m_URLComponents.lpszHostName,
		m_URLComponents.nPort,
		NULL,
		NULL,
		dwService,
		0,
		(m_bAsync == VARIANT_TRUE) ? (DWORD_PTR) this : (DWORD_PTR) NULL);
#ifdef _DEBUG
	OutputDebugFormat(L"m_hConnect = 0x%p\r\n", m_hConnect);
#endif
	if (m_hConnect == NULL)
	{
		dwLastError = GetLastError();
		if (dwLastError != ERROR_IO_PENDING)
		{
			hr = HRESULT_FROM_WIN32(dwLastError);
			return hr;
		}
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoHttpOpenRequest()
{
	HRESULT hr = S_OK;
	DWORD dwLastError = ERROR_SUCCESS;

	CComBSTR bstrObjectName;
	bstrObjectName.Append(m_URLComponents.lpszUrlPath);
	bstrObjectName.Append(m_URLComponents.lpszExtraInfo);

	DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE;
	if (m_URLComponents.nScheme == INTERNET_SCHEME_HTTPS)
	{
#ifdef UNDER_CE
		dwFlags |= INTERNET_FLAG_IGNORE_CERT_CN_INVALID;
		dwFlags |= INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;
#endif
		dwFlags |= INTERNET_FLAG_SECURE;
	}
	dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;

	m_state = stateHttpOpenRequest;
	m_hRequest = HttpOpenRequest(
		m_hConnect,
		m_bstrMethod,
		bstrObjectName,
		NULL,
		NULL,
		NULL,
		dwFlags,
		(DWORD_PTR) this);
#ifdef _DEBUG
	OutputDebugFormat(L"m_hRequest = 0x%p\r\n", m_hRequest);
#endif
	if (m_hRequest == NULL)
	{
		dwLastError = GetLastError();
		if (dwLastError != ERROR_IO_PENDING)
		{
			hr = HRESULT_FROM_WIN32(dwLastError);
			return hr;
		}
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoSend(VARIANT& varData)
{
	VARIANT* pvarData = &varData;
	if (pvarData->vt == (VT_VARIANT | VT_BYREF))
	{
		pvarData = pvarData->pvarVal;
	}
	switch (pvarData->vt)
	{
	case VT_EMPTY:
	case VT_NULL:
		return DoHttpSendRequest(0, 0);

	case VT_BSTR:
		return DoSend(V_BSTR(pvarData));
	}

	OutputDebugFormat(L"Error: DoHttpSendRequest (vt=%d) not implemented\r\n", pvarData->vt);
	return E_NOTIMPL;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoSend(BSTR bstrData)
{
	int nLenW = SysStringLen(bstrData);
	if (nLenW == 0)
	{
		return DoHttpSendRequest(0, 0);
	}

	int nLenA = WideCharToMultiByte(CP_UTF8, 0, bstrData, nLenW + 1, NULL, 0, NULL, NULL);
	if (nLenA <= 1)
	{
		return DoHttpSendRequest(0, 0);
	}

	CComHeapPtr<BYTE> spDataA;
	if (!spDataA.AllocateBytes(nLenA))
	{
		return E_OUTOFMEMORY;
	}
	ZeroMemory((BYTE*) spDataA, nLenA);
	WideCharToMultiByte(CP_UTF8, 0, bstrData, nLenW + 1, (LPSTR) (BYTE*) spDataA, nLenA, NULL, NULL);
	return DoHttpSendRequest((BYTE*) spDataA, nLenA - 1);
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoHttpSendRequest(LPBYTE pData, int nData)
{
	HRESULT hr = S_OK;
	DWORD dwLastError = ERROR_SUCCESS;
	BOOL bOk = TRUE;

	WCHAR szHeaders[] = L"Content-Type: application/x-www-form-urlencoded\n";
	int nHeaders = wcslen(szHeaders);

	m_state = stateHttpSendRequest;
	bOk = HttpSendRequest(m_hRequest, szHeaders, nHeaders, pData, nData);
	CHECKHR(put_ReadyState(2));
	if (bOk != TRUE)
	{
		dwLastError = GetLastError();
		if (dwLastError != ERROR_IO_PENDING)
		{
#ifdef _DEBUG
			OutputDebugFormat(L"HttpSendRequest -> FALSE (GetLastError: %d)\r\n", dwLastError);
#endif
			hr = HRESULT_FROM_WIN32(dwLastError);
			return hr;
		}
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::DoInternetReadFile()
{
	HRESULT hr = S_OK;
	BOOL bOk = TRUE;
	DWORD dwLastError;

	m_state = stateInternetReadFile;

	BYTE pData[1024] = { };
	DWORD dwRead = 0;
	while (bOk == TRUE)
	{
		dwRead = 0;
		bOk = ::InternetReadFile(m_hRequest, pData, 1024, &dwRead);
		if (bOk == TRUE)
		{
#ifdef _DEBUG
			OutputDebugFormat(L"InternetReadFile -> TRUE (%d bytes)\r\n", dwRead);
#endif
			if (dwRead == 0)
			{
				if (m_ResponseFile)
				{
					CHECKHR(m_ResponseFile->Close());
					m_ResponseFile = NULL;
				}
				CHECKHR(put_ReadyState(4));
				if (m_bAsync == VARIANT_TRUE)
				{
					//Release();
				}
				return S_OK;
			}

			OnContentDownload(pData, dwRead);
		}
	}
	dwLastError = GetLastError();
	if (dwLastError != ERROR_IO_PENDING)
	{
		OutputDebugFormat(L"InternetReadFile -> FALSE (GetLastError:%d)\r\n", dwLastError);
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::OnContentDownload(BYTE* pContent, LONG nContent)
{
	HRESULT hr = S_OK;

	if (nContent <= 0)
	{
		return S_OK;
	}

	if (m_bstrResponsePath.Length() > 0)
	{
		if (m_ResponseFile == NULL)
		{
			CComPtr<IQSFile> spIQSFile;
			VARIANT_BOOL bOk = VARIANT_TRUE;
			CHECKHR(CQSFile::_CreatorClass::CreateInstance(NULL, IID_IQSFile, (void**) &spIQSFile));
			CHECKHR(spIQSFile->Open(m_bstrResponsePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, &bOk));
			if (bOk != VARIANT_TRUE) return S_FALSE;
			CHECKHR(spIQSFile->QueryInterface(IID_IQSFile, (void**) &m_ResponseFile));
		}

		CComPtr<IStream> spIStream;
		CHECKHR(m_ResponseFile->QueryInterface(IID_IStream, (void**) &spIStream));
		DWORD dwWritten = 0;
		hr = spIStream->Write(pContent, nContent, &dwWritten);
		if (FAILED(hr))
		{
			return hr;
		}
		return hr;
	}

	ULONG dwCount = m_ResponseBody.GetCount();
	m_ResponseBody.Resize(dwCount + nContent);
	memcpy(&m_ResponseBody[(LONG) dwCount], pContent, nContent);

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::OnReadyStateChange()
{
	HRESULT hr = S_OK;
	if (!m_OnReadyStateChange) return S_OK;
	if (m_bAsync == VARIANT_TRUE)
	{
		CComPtr<IPropertyBag> spIPropertyBag;
		hr = m_OnReadyStateChange->QueryInterface(IID_IPropertyBag, (void**) &spIPropertyBag);
		if (SUCCEEDED(hr) && spIPropertyBag)
		{
			CComVariant varhWnd;
			hr = spIPropertyBag->Read(OLESTR("hWnd"), &varhWnd, NULL);
			if (SUCCEEDED(hr) && varhWnd.vt == VT_I4)
			{
				HWND hWnd = (HWND) V_I4(&varhWnd);
				if (hWnd != NULL)
				{
					PostMessage(hWnd, WM_USER, 0, (LPARAM) (IDispatch*) m_OnReadyStateChange);
				}
			}
		}
	}
	else
	{
		CComVariant varResult;
		CHECKHR(InvokeMethod(m_OnReadyStateChange, NULL, CComVariant(), CComVariant(), CComVariant(), &varResult));
	}
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::Open(BSTR bstrMethod, BSTR bstrURL, VARIANT varAsync)
{
	HRESULT hr = S_OK;
	//VARIANT_BOOL bIsOnline = VARIANT_FALSE;
	//hr = get_IsOnline(&bIsOnline);
	DWORD dwLastError = ERROR_SUCCESS;
	BOOL bOk = FALSE;

	if (m_hInternet || m_hConnect || m_hRequest || m_hOpenEvent)
	{
		Close();
	}

	bOk = InternetCheckConnection(bstrURL, FLAG_ICC_FORCE_CONNECTION, 0);
	if (bOk != TRUE)
	{
		dwLastError = GetLastError();
#ifdef _DEBUG
		OutputDebugFormat(L"InternetCheckConnection -> FALSE (GetLastError: %d)\r\n", dwLastError);
#endif
		return S_FALSE;
	}

	m_bstrMethod = bstrMethod;
	m_bstrURL = bstrURL;
	CHECKHR(VariantToBOOL(varAsync, &m_bAsync, VARIANT_TRUE));
	//CHECKHR(m_ResponseBody.Resize((ULONG) 0));

	if (m_bAsync == VARIANT_TRUE)
	{
		m_hOpenEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
		CHECKHR(DoInternetOpen());
		CHECKHR(DoInternetConnect());
		DWORD dwTick = GetTickCount();
		//AddRef();
		DWORD dwWait = WaitForSingleObject(m_hOpenEvent, m_nOpenTimeout >= 0 ? (DWORD) m_nOpenTimeout : INFINITE);
		if (dwWait != WAIT_OBJECT_0)
		{
#ifdef _DEBUG
			OutputDebugFormat(L"WaitForSingleObject(m_hOpenEvent) != WAIT_OBJECT_0 (%d)\r\n", dwWait);
#endif
		}
		DWORD dwElapse = GetTickCount() - dwTick;
		OutputDebugFormat(L"Waited for %d ms\r\n", dwElapse);
		bOk = CloseHandle(m_hOpenEvent);
		if (bOk != TRUE)
		{
#ifdef _DEBUG
			dwLastError = GetLastError();
			OutputDebugFormat(L"CloseHandle(m_hOpenEvent) = FALSE (GetLastError:%d)\r\n", dwLastError);
#endif
			//Release();
		}
		else
		{
			m_hOpenEvent = NULL;
			CHECKHR(put_ReadyState(1));
		}
	}
	else
	{
		CHECKHR(DoInternetOpen());
		CHECKHR(DoInternetConnect());
		CHECKHR(DoHttpOpenRequest());
		CHECKHR(put_ReadyState(1));
	}

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSNet::Send(VARIANT varBody)
{
	HRESULT hr = S_OK;
	CHECKHR(DoSend(varBody));
	if (m_bAsync != VARIANT_TRUE)
	{
		CHECKHR(put_ReadyState(3));
		CHECKHR(DoInternetReadFile());
	}
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

void CALLBACK CQSNet::InternetStatusCallback(
	HINTERNET hInternet,
	DWORD_PTR dwContext,
	DWORD dwInternetStatus,
	LPVOID lpvStatusInformation,
	DWORD dwStatusInformationLength)
{
	((CQSNet*) dwContext)->InternetStatusCallback(hInternet, dwInternetStatus, lpvStatusInformation, dwStatusInformationLength);
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

void CQSNet::InternetStatusCallback(
	HINTERNET hInternet,
	DWORD dwInternetStatus,
	LPVOID lpvStatusInformation,
	DWORD dwStatusInformationLength)
{
	switch (dwInternetStatus)
	{
	case INTERNET_STATUS_HANDLE_CREATED:
		switch (m_state)
		{
		case stateInternetConnect:
			m_hConnect = *(HINTERNET*) lpvStatusInformation;
#ifdef _DEBUG
			OutputDebugFormat(L"m_hConnect (async) = 0x%p\r\n", m_hConnect);
#endif
			DoHttpOpenRequest();
			//DoInternetOpen();
			break;

		case stateHttpOpenRequest:
			m_hRequest = *(HINTERNET*) lpvStatusInformation;
#ifdef _DEBUG
			OutputDebugFormat(L"m_hRequest (async) = 0x%p\r\n", m_hRequest);
#endif
			SetEvent(m_hOpenEvent);
			//DoHttpSendRequest();
			break;
		}
		break;

	case INTERNET_STATUS_REQUEST_COMPLETE:
		switch (m_state)
		{
		case stateHttpSendRequest:
			//OnHttpSendRequestComplete(lpvStatusInformation);
			put_ReadyState(3);
			OutputDebugRequest(m_hRequest);
			DoInternetReadFile();
			break;

		//case asyncStateInternetReadFileEx:
			/*
			m_asyncState = asyncStateUnknown;
			SetEvent(m_hRequestCompleteEvent);
			*/
			//break;

		}

		break;

	}
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
