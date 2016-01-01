// QSGlobal.h : Declaration of the CQSGlobal

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



// CQSGlobal

class ATL_NO_VTABLE CQSGlobal :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQSGlobal, &CLSID_QSGlobal>,
	public ISupportErrorInfo,
	public IPropertyBag,
	public IDispatchImpl<IQSGlobal, &IID_IQSGlobal, &LIBID_QSLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CQSGlobal() : m_hWnd(NULL)
	{
	}

#ifndef _CE_DCOM
DECLARE_REGISTRY_RESOURCEID(IDR_QSGLOBAL)
#endif


BEGIN_COM_MAP(CQSGlobal)
	COM_INTERFACE_ENTRY(IQSGlobal)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IPropertyBag)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPropertyBag
	STDMETHOD(Read)(LPCOLESTR pszPropName, VARIANT *pVar, IErrorLog *pErrorLog);
	STDMETHOD(Write)(LPCOLESTR pszPropName, VARIANT *pVar);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	CComPtr<IActiveScript> m_spIActiveScript;
	HWND m_hWnd;

public:
	STDMETHOD(GetRef)(BSTR bstrName, IDispatch** ppIDispatch);

};

OBJECT_ENTRY_AUTO(__uuidof(QSGlobal), CQSGlobal)
