#include "stdafx.h"
#include "OutputDebugMemoryStatus.h"

void OutputDebugMemoryStatus()
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
	_stprintf(szText, _T("Phys=%d (%d) Virt=%d (%d)"),
		MS.dwAvailPhys, MS.dwAvailPhys - LastMS.dwAvailPhys,
		MS.dwAvailVirtual, MS.dwAvailVirtual - LastMS.dwAvailVirtual);
	OutputDebugString(szText);

	LastMS = MS;
}
