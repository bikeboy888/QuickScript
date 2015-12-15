// QuickScriptDlg.cpp : implementation file
//

#include "stdafx.h"
#include "QuickScript.h"
#include "QuickScriptDlg.h"
#include "..\QS\QS_i.c"
#include "..\QSUtil\RegSvr.h"

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
		edtScript->SetWindowText(L"Function f(x)\r\n  f = x + 2\r\nEnd Function\r\n");
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

	if (!m_spIQSScriptSite)
	{
		hr = m_spIQSScriptSite.CoCreateInstance(CLSID_QSScriptSite);
		if (FAILED(hr))
		{
			RegSvr(L"QS.dll");
			hr = m_spIQSScriptSite.CoCreateInstance(CLSID_QSScriptSite);
		}
		hr = m_spIQSScriptSite->put_ScriptEngine(CComBSTR(L"VBScript"));
	}

	CEdit* edtScript = (CEdit*) GetDlgItem(EDT_SCRIPT);
	if (!edtScript) return;
	int nLen = edtScript->GetWindowTextLength();
	CComBSTR bstrScript(nLen, (LPCOLESTR) NULL);
	edtScript->GetWindowText((BSTR) bstrScript, nLen + 1);

	VARIANT varContext = { };
	CComVariant varResult;
	CComBSTR bstrID(L"Script1234");
	//hr = m_spIQSScriptSite->ImportScript(CComBSTR(L"function f(x) { return x*x + 1; }"), CComBSTR(L"JScript"), CComVariant(bstrID));
	//hr = m_spIQSScriptSite->ImportScript(CComBSTR(L"Function f(x)\r\n  f = x*x + 1\r\nEnd Function\r\n"), CComBSTR(L"VBScript"), CComVariant(bstrID));
	hr = m_spIQSScriptSite->Execute(bstrScript, varContext);


	hr = m_spIQSScriptSite->InvokeMethod(CComBSTR(L"f"), CComVariant(5), CComVariant(), CComVariant(), &varResult);
//m_spIQSScriptSite->ParsePinni

	CComVariant varResultBSTR;
	varResultBSTR.ChangeType(VT_BSTR, &varResult);
	TCHAR szText[1024] = { };
	_stprintf(szText, _T("Result: %s\r\n"), V_BSTR(&varResultBSTR));
	OutputDebugString(szText);
	MessageBox(szText, _T("Output"), MB_OK);
}

void CQuickScriptDlg::OnBnClickedCleanup()
{
	HRESULT hr = S_OK;

	if (m_spIQSScriptSite)
	{
		hr = m_spIQSScriptSite->Close();
		m_spIQSScriptSite = NULL;
	}

	CoFreeUnusedLibrariesEx(0, 0);
	CoFreeUnusedLibrariesEx(0, 0);
}
