// QuickScriptDlg.cpp : implementation file
//

#include "stdafx.h"
#include "QuickScript.h"
#include "QuickScriptDlg.h"
#include "..\QS\QS_i.h"
#include "..\QS\QS_i.c"

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
#if defined(_DEVICE_RESOLUTION_AWARE) && !defined(WIN32_PLATFORM_WFSP)
	ON_WM_SIZE()
#endif
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(BTN_RUN, &CQuickScriptDlg::OnBnClickedRun)
END_MESSAGE_MAP()


// CQuickScriptDlg message handlers

BOOL CQuickScriptDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

#if defined(_DEVICE_RESOLUTION_AWARE) && !defined(WIN32_PLATFORM_WFSP)
void CQuickScriptDlg::OnSize(UINT /*nType*/, int /*cx*/, int /*cy*/)
{
	if (AfxIsDRAEnabled())
	{
		DRA::RelayoutDialog(
			AfxGetResourceHandle(), 
			this->m_hWnd, 
			DRA::GetDisplayMode() != DRA::Portrait ? 
			MAKEINTRESOURCE(IDD_QUICKSCRIPT_DIALOG_WIDE) : 
			MAKEINTRESOURCE(IDD_QUICKSCRIPT_DIALOG));
	}
}
#endif


void CQuickScriptDlg::OnBnClickedRun()
{
	HRESULT hr = S_OK;

	CComPtr<IQSScriptSite> spScriptSite;
	hr = spScriptSite.CoCreateInstance(CLSID_QSScriptSite);
	hr = spScriptSite->put_ScriptEngine(CComBSTR(L"VBScript"));
	spScriptSite = NULL;

	// TODO: Add your control notification handler code here
}
