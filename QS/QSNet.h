// QSNet.h : Declaration of the CQSNet

#pragma once
#ifdef STANDARDSHELL_UI_MODEL
#include "resource.h"
#endif
#ifdef POCKETPC2003_UI_MODEL
#include "resourceppc.h"
#endif
#ifdef SMARTPHONE2003_UI_MODEL
#include "resourcesp.h"
#endif
#ifdef AYGSHELL_UI_MODEL
#include "resourceayg.h"
#endif

#include "QS_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CQSNet

class ATL_NO_VTABLE CQSNet :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQSNet, &CLSID_QSNet>,
	public ISupportErrorInfo,
	public IDispatchImpl<IQSNet, &IID_IQSNet, &LIBID_QSLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CQSNet();

#ifndef _CE_DCOM
DECLARE_REGISTRY_RESOURCEID(IDR_QSNET)
#endif

BEGIN_COM_MAP(CQSNet)
	COM_INTERFACE_ENTRY(IQSNet)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
		if (m_hInternet || m_hConnect || m_hRequest || m_hOpenEvent)
		{
			Close();
		}
	}

protected:
	enum State
	{
		stateUnknown,
		stateInternetOpen,
		stateInternetConnect,
		stateHttpOpenRequest,
		stateHttpSendRequest,
		stateInternetReadFile
	};

protected:
	CComBSTR m_bstrMethod;
	CComBSTR m_bstrURL;
	VARIANT_BOOL m_bAsync;
	URL_COMPONENTS m_URLComponents;
	TCHAR m_szScheme[1024];
	TCHAR m_szHostName[1024];
	TCHAR m_szUserName[1024];
	TCHAR m_szPassword[1024];
	TCHAR m_szUrlPath[1024];
	TCHAR m_szExtraInfo[1024];
	HINTERNET m_hInternet;
	HINTERNET m_hConnect;
	HINTERNET m_hRequest;
	HANDLE m_hOpenEvent;
	LONG m_nOpenTimeout;
	enum State m_state;
	CComSafeArray<BYTE> m_ResponseBody;
	CComBSTR m_bstrResponsePath;
	CComPtr<IQSFile> m_ResponseFile;
	CComPtr<IDispatch> m_OnReadyStateChange;
	LONG m_nReadyState;

	STDMETHOD(DoInternetOpen)();
	STDMETHOD(DoInternetConnect)();
	STDMETHOD(DoHttpOpenRequest)();
	STDMETHOD(DoHttpSendRequest)();
	STDMETHOD(DoInternetReadFile)();
	STDMETHOD(OnContentDownload)(BYTE* pContent, LONG nContent);
	STDMETHOD(OnReadyStateChange)();

protected:
	static void CALLBACK InternetStatusCallback(HINTERNET, DWORD_PTR, DWORD, LPVOID, DWORD);
	void InternetStatusCallback(HINTERNET, DWORD, LPVOID, DWORD);

public:
	STDMETHOD(get_OnReadyStateChange)(VARIANT* pvarDispatch);
	STDMETHOD(putref_OnReadyStateChange)(VARIANT varDispatch);
	STDMETHOD(put_OnReadyStateChange)(VARIANT varDispatch);
	STDMETHOD(get_OpenTimeout)(LONG* pnOpenTimeout);
	STDMETHOD(put_OpenTimeout)(LONG nOpenTimeout);
	STDMETHOD(get_ReadyState)(LONG* pnReadyState);
	STDMETHOD(put_ReadyState)(LONG nReadyState);
	STDMETHOD(get_ResponsePath)(BSTR* pbstrPath);
	STDMETHOD(put_ResponsePath)(BSTR bstrPath);
	STDMETHOD(get_ResponseText)(BSTR* pbstrText);
	STDMETHOD(get_Status)(LONG* pnStatus);
	STDMETHOD(Close)();
	STDMETHOD(Open)(BSTR bstrMethod, BSTR bstrURL, VARIANT varAsync);
	STDMETHOD(Send)(VARIANT varBody);

};

OBJECT_ENTRY_AUTO(__uuidof(QSNet), CQSNet)
