// QSJScript.h : Declaration of the CQSJScript

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



// CQSJScript

class ATL_NO_VTABLE CQSJScript :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQSJScript, &CLSID_QSJScript>,
	public ISupportErrorInfo,
	public IDispatchImpl<IQSJScript, &IID_IQSJScript, &LIBID_QSLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CQSJScript()
	{
	}

#ifndef _CE_DCOM
DECLARE_REGISTRY_RESOURCEID(IDR_QSJSCRIPT)
#endif


BEGIN_COM_MAP(CQSJScript)
	COM_INTERFACE_ENTRY(IQSJScript)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();

	void FinalRelease()
	{
		if (m_spIQSScriptSite)
		{
			m_spIQSScriptSite->Close();
			m_spIQSScriptSite = NULL;
		}
		CoFreeUnusedLibrariesEx(0, 0);
		CoFreeUnusedLibrariesEx(0, 0);
	}

public:
	CComPtr<IQSScriptSite> m_spIQSScriptSite;

public:
	STDMETHOD(EncodeURIComponent)(BSTR bstrURI, BSTR* pbstrEncodedURI);

};

OBJECT_ENTRY_AUTO(__uuidof(QSJScript), CQSJScript)
