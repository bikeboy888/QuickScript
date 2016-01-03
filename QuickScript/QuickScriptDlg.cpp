// QuickScriptDlg.cpp : implementation file
//

#include "stdafx.h"
#include "QuickScript.h"
#include "QuickScriptDlg.h"
#include "..\QS\QS_i.c"
#include "..\QSUtil\RegSvr.h"
#include "..\QSUtil\OutputDebugMemoryStatus.h"
#include "..\QSUtil\OutputDebugVariant.h"
#include "..\QSUtil\OutputDebugFormat.h"
#include "..\QSUtil\InvokeMethod.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CQuickScriptDlg dialog

CQuickScriptDlg::CQuickScriptDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CQuickScriptDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CQuickScriptDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CQuickScriptDlg, CDialog)
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(BTN_RUN, &CQuickScriptDlg::OnBnClickedRun)
	ON_BN_CLICKED(BTN_CLEANUP, &CQuickScriptDlg::OnBnClickedCleanup)
	ON_MESSAGE(WM_USER, &CQuickScriptDlg::OnScriptInvoke)
END_MESSAGE_MAP()


// CQuickScriptDlg message handlers

BOOL CQuickScriptDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	CEdit* edtScript = (CEdit*) GetDlgItem(EDT_SCRIPT);
	if (edtScript)
	{
		edtScript->SetWindowText(
			L"Dim net, body, a, bigdata\r\n"
			L"\r\n"
			L"Function rsc\r\n"
			L"  a = 5\r\n"
			L"  If net.ReadyState = 4 Then\r\n"
			L"    MsgBox net.responseText\r\n"
			L"  End If\r\n"
			L"End Function\r\n"
			L"\r\n"
			L"  Set bigdata = CreateObject(\"QS.File\")\r\n"
			L"  bigdata.Open \"\\Storage Card\\json.txt\", &h80000000, &h3, 3\r\n"
			L"  REM MsgBox bigdata.ReadAll\r\n"
			L"\r\n"
			L"  Set net = CreateObject(\"QS.Net\")\r\n"
			L"  Set net.OnReadyStateChange = GetRef(\"rsc\")\r\n"
			L"  REM net.Open \"GET\", \"http://www.arcgis.com/sharing/rest?f=json\", True\r\n"
			L"  REM net.Open \"POST\", \"https://www.arcgis.com/sharing/rest/generateToken\", True\r\n"
			L"  net.Open \"POST\", \"http://www.arcgis.com/sharing/rest/info\", True\r\n"
			L"  body = \"f=json\"\r\n"
			L"  REM net.Send body\r\n"
			L"  net.Send bigdata\r\n"
			L"\r\n"
			L"Function f(x)\r\n"
			L"  f = x + 2\r\n"
			L"End Function\r\n"
			);
	}

	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CQuickScriptDlg::OnSize(UINT /*nType*/, int /*cx*/, int /*cy*/)
{
	/*
	if (AfxIsDRAEnabled())
	{
		DRA::RelayoutDialog(
			AfxGetResourceHandle(), 
			this->m_hWnd, 
			DRA::GetDisplayMode() != DRA::Portrait ? 
			MAKEINTRESOURCE(IDD_QUICKSCRIPT_DIALOG_WIDE) : 
			MAKEINTRESOURCE(IDD_QUICKSCRIPT_DIALOG));
	}
	*/
	LayoutControls();
}

HRESULT CQuickScriptDlg::LayoutControls()
{
	if (!m_hWnd) return S_FALSE;

	CButton* btnRun = (CButton*) GetDlgItem(BTN_RUN);
	if (!btnRun) return S_FALSE;
	CButton* btnCleanUp = (CButton*) GetDlgItem(BTN_CLEANUP);
	if (!btnCleanUp) return S_FALSE;
	CEdit* edtScript = (CEdit*) GetDlgItem(EDT_SCRIPT);
	if (!edtScript) return S_FALSE;

	RECT rc = { };
	GetClientRect(&rc);
	LONG w = rc.right - rc.left;
	LONG h = rc.bottom - rc.top;
	LONG m = 5;

	RECT rcRun = { };
	btnRun->GetClientRect(&rcRun);
	LONG wRun = rcRun.right - rcRun.left;
	LONG hRun = rcRun.bottom - rcRun.top;
	rcRun.left = m;
	rcRun.right = rcRun.left + wRun;
	rcRun.top = m;
	rcRun.bottom = rcRun.top + hRun;
	btnRun->MoveWindow(rcRun.left, rcRun.top, rcRun.right - rcRun.left, rcRun.bottom - rcRun.top);

	RECT rcCleanUp = { };
	btnCleanUp->GetClientRect(&rcCleanUp);
	LONG wCleanUp = rcCleanUp.right - rcCleanUp.left;
	LONG hCleanUp = rcCleanUp.bottom - rcCleanUp.top;
	rcCleanUp.left = rcRun.right + m;
	rcCleanUp.right = rcCleanUp.left + wCleanUp;
	rcCleanUp.top = m;
	rcCleanUp.bottom = rcCleanUp.top + hCleanUp;
	btnCleanUp->MoveWindow(rcCleanUp.left, rcCleanUp.top, rcCleanUp.right - rcCleanUp.left, rcCleanUp.bottom - rcCleanUp.top);

	RECT rcScript = { };
	rcScript.left = m;
	rcScript.right = w - m;
	rcScript.top = rcRun.bottom + m;
	rcScript.bottom = h - m;
	edtScript->MoveWindow(rcScript.left, rcScript.top, rcScript.right - rcScript.left, rcScript.bottom - rcScript.top);

	return S_OK;
}


void CQuickScriptDlg::OnBnClickedRun()
{
	HRESULT hr = S_OK;

	if (!m_spIQSApplication)
	{
		hr = m_spIQSApplication.CoCreateInstance(CLSID_QSApplication);
		if (FAILED(hr))
		{
			RegSvr(L"QS.dll");
			hr = m_spIQSApplication.CoCreateInstance(CLSID_QSApplication);
		}
	}

	hr = m_spIQSApplication->put_UserProperties(CComBSTR(L"pi"), CComVariant(3.1415));
	CComVariant pi;
	hr = m_spIQSApplication->get_UserProperties(CComBSTR(L"pi"), &pi);

	if (!m_spIXMLDOMDocument)
	{
		hr = m_spIXMLDOMDocument.CoCreateInstance(OLESTR("Microsoft.XMLDOM"));
		VARIANT_BOOL bSuccessful = VARIANT_FALSE;
		hr = m_spIXMLDOMDocument->loadXML(CComBSTR(L"<HELLO><WORLD/></HELLO>"), &bSuccessful);
	}

	if (!m_spIQSScriptSite)
	{
		//hr = E_FAIL;
		hr = m_spIQSScriptSite.CoCreateInstance(CLSID_QSScriptSite);
		hr = m_spIQSScriptSite->put_ScriptEngine(CComBSTR(L"VBScript"));
		hr = m_spIQSScriptSite->SetItem(CComBSTR("Prefs"), SCRIPTITEM_ISVISIBLE, m_spIXMLDOMDocument);
	}

	hr = m_spIQSScriptSite->put_hWnd((OLE_HANDLE) m_hWnd);

	//CComBSTR bstrURL(L"http://www.arcgis.com/sharing/rest?f=pjson");
	//CComBSTR bstrURL(L"http://tviview.abc.net.au/iview/api2/?keyword=a-l");
	CComBSTR bstrURL(L"http://www.arcgis.com/sharing/rest/info?f=pjson");
	m_spIQSNet = NULL;
	/*
	hr = m_spIQSNet.CoCreateInstance(CLSID_QSNet);
	//hr = m_spIQSNet->Open(CComBSTR("GET"), bstrURL, CComVariant(true));
	hr = m_spIQSNet->put_ResponsePath(CComBSTR(L"\\My Documents\\Output.txt"));
	hr = m_spIQSNet->Open(CComBSTR("GET"), bstrURL, CComVariant(false));
	hr = m_spIQSNet->Send(CComVariant());
	LONG nStatus = 0;
	hr = m_spIQSNet->get_Status(&nStatus);
	OutputDebugFormat(L"Status = %d\r\n", nStatus);
	LONG nOpenTimeout = 0;
	hr = m_spIQSNet->get_OpenTimeout(&nOpenTimeout);
	OutputDebugFormat(L"OpenTimeout = %d\r\n", nOpenTimeout);
	nOpenTimeout = 5000;
	hr = m_spIQSNet->put_OpenTimeout(nOpenTimeout);
	hr = m_spIQSNet->get_OpenTimeout(&nOpenTimeout);
	OutputDebugFormat(L"OpenTimeout = %d\r\n", nOpenTimeout);
	CComBSTR bstrText;
	m_spIQSNet->get_ResponseText(&bstrText);
	OutputDebugFormat(L"Text = %s\r\n", (BSTR) bstrText);
	*/

	//hr = spIQSNet->Close();
	//spIQSNet = NULL;

	/*
	CComPtr<IQSJScript> spIQSJScript;
	hr = spIQSJScript.CoCreateInstance(CLSID_QSJScript);
	spIQSJScript = NULL;
	*/

	CEdit* edtScript = (CEdit*) GetDlgItem(EDT_SCRIPT);
	if (!edtScript) return;
	int nLen = edtScript->GetWindowTextLength();
	CComBSTR bstrScript(nLen, (LPCOLESTR) NULL);
	int nLen2 = bstrScript.Length();
	edtScript->GetWindowText((BSTR) bstrScript, nLen + 1);

	VARIANT varContext = { };
	CComVariant varResult;
	CComBSTR bstrID(L"Script1234");
	//hr = m_spIQSScriptSite->ImportScript(CComBSTR(L"function f(x) { return x*x + 1; }"), CComBSTR(L"JScript"), CComVariant(bstrID));
	//hr = m_spIQSScriptSite->ImportScript(CComBSTR(L"Function f(x)\r\n  f = x*x + 1\r\nEnd Function\r\n"), CComBSTR(L"VBScript"), CComVariant(bstrID));
	hr = m_spIQSScriptSite->Execute(bstrScript, varContext);

	hr = m_spIQSScriptSite->InvokeMethod(CComBSTR(L"f"), CComVariant(5), CComVariant(), CComVariant(), &varResult);
	OutputDebugString(L"Result: ");
	OutputDebugVariant(varResult);
	OutputDebugString(L"\r\n");

	OutputDebugMemoryStatus();
	OutputDebugString(L"\r\n");
}

void CQuickScriptDlg::OnBnClickedCleanup()
{
	HRESULT hr = S_OK;

	if (m_spIQSApplication)
	{
		m_spIQSApplication = NULL;
	}

	if (m_spIQSNet)
	{
		m_spIQSNet = NULL;
	}

	if (m_spIXMLDOMDocument)
	{
		m_spIXMLDOMDocument = NULL;
	}

	if (m_spIQSScriptSite)
	{
		hr = m_spIQSScriptSite->Close();
		m_spIQSScriptSite = NULL;
	}

	CoFreeUnusedLibrariesEx(0, 0);
	CoFreeUnusedLibrariesEx(0, 0);

	OutputDebugMemoryStatus();
	OutputDebugString(L"\r\n");
}

LRESULT CQuickScriptDlg::OnScriptInvoke(WPARAM wParam, LPARAM lParam)
{
	IDispatch* pIDispatch = (IDispatch*) lParam;
	CComVariant varResult;
	HRESULT hr = InvokeMethod(pIDispatch, NULL, CComVariant(), CComVariant(), CComVariant(), &varResult);
	return 0;
}

