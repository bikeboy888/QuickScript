// QuickScriptDlg.h : header file
//

#pragma once

#include "..\QS\QS_i.h"

// CQuickScriptDlg dialog
class CQuickScriptDlg : public CDialog
{
// Construction
public:
	CQuickScriptDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_QUICKSCRIPT_DIALOG };


	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

// Implementation
protected:
	HICON m_hIcon;
	CComPtr<IQSScriptSite> m_spIQSScriptSite;
	CComPtr<IQSNet> m_spIQSNet;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT /*nType*/, int /*cx*/, int /*cy*/);
	HRESULT LayoutControls();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRun();
	afx_msg void OnBnClickedCleanup();
};
