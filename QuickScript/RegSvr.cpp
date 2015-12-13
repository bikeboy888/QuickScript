#include "stdafx.h"
#include "RegSvr.h"

HRESULT RegSvr(LPCWSTR szPath)
{
	HRESULT hr = S_OK;

	HMODULE hModule = LoadLibrary(szPath);
	if (hModule == NULL)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	HRESULT (__stdcall *DllRegisterServer)();
	(FARPROC&) DllRegisterServer = GetProcAddress(hModule, _T("DllRegisterServer"));
	if (DllRegisterServer == NULL)
	{
		FreeLibrary(hModule);
		return E_NOTIMPL;
	}

	hr = DllRegisterServer();
	if (FAILED(hr))
	{
		FreeLibrary(hModule);
		return hr;
	}

	BOOL bOk = FreeLibrary(hModule);
	if (bOk != TRUE)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	return S_OK;
}

