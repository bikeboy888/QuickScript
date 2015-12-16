#include "stdafx.h"
#include "DebugMemoryStatus.h"

void DebugMemoryStatus()
{
	TCHAR szText[1024] = { };
	static MEMORYSTATUS LastMS = { };
	MEMORYSTATUS MS = { };

	::GlobalMemoryStatus(&MS);
	/*
	DWORD  dwLength;
  DWORD  dwMemoryLoad;
  SIZE_T dwTotalPhys;
  SIZE_T dwAvailPhys;
  SIZE_T dwTotalPageFile;
  SIZE_T dwAvailPageFile;
  SIZE_T dwTotalVirtual;
  SIZE_T dwAvailVirtual;
  */
	_stprintf(szText, _T("Phys=%d (%d) Virt=%d (%d)\r\n"),
		MS.dwAvailPhys, MS.dwAvailPhys - LastMS.dwAvailPhys,
		MS.dwAvailVirtual, MS.dwAvailVirtual - LastMS.dwAvailVirtual);
	OutputDebugString(szText);

	LastMS = MS;
}
