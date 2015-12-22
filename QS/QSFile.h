// QSFile.h : Declaration of the CQSFile

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



// CQSFile

class ATL_NO_VTABLE CQSFile :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQSFile, &CLSID_QSFile>,
	public ISupportErrorInfo,
	public IStream,
	public IDispatchImpl<IQSFile, &IID_IQSFile, &LIBID_QSLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CQSFile() : m_hFile(INVALID_HANDLE_VALUE), m_CodePage(0)
	{
	}

#ifndef _CE_DCOM
DECLARE_REGISTRY_RESOURCEID(IDR_QSFILE)
#endif


BEGIN_COM_MAP(CQSFile)
	COM_INTERFACE_ENTRY(IQSFile)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IStream)
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
		if (m_hFile != INVALID_HANDLE_VALUE)
		{
			Close();
		}
	}

protected:
	HANDLE m_hFile;
	UINT m_CodePage;

public:
	STDMETHOD(Close)();
	STDMETHOD(Open)(BSTR bstrPath, LONG nDesiredAccess, LONG nShareMode, LONG nCreationDisposition, VARIANT_BOOL* pbOk);
	STDMETHOD(ReadAll)(BSTR* pbstrText);

	// ISequentialStream
	STDMETHOD(Read)(void *pv, ULONG cb, ULONG *pcbRead);
	STDMETHOD(Write)(const void *pv, ULONG cb, ULONG *pcbWritten);

	// IStream
	STDMETHOD(Seek)(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition);
	STDMETHOD(SetSize)(ULARGE_INTEGER libNewSize);
	STDMETHOD(CopyTo)(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten);
	STDMETHOD(Commit)(DWORD grfCommitFlags);
	STDMETHOD(Revert)( void) { return E_NOTIMPL; }
	STDMETHOD(LockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) { return E_NOTIMPL; }
	STDMETHOD(UnlockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) { return E_NOTIMPL; }
	STDMETHOD(Stat)(STATSTG *pstatstg, DWORD grfStatFlag);
	STDMETHOD(Clone)(IStream **ppstm);

};

OBJECT_ENTRY_AUTO(__uuidof(QSFile), CQSFile)
