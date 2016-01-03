#include "stdafx.h"
#include "OutputDebugRequest.h"
#include "OutputDebugFormat.h"

void OutputDebugRequest(HINTERNET hRequest)
{
	DWORD dwSize = 4;
	DWORD dwStatus = 0;
	BOOL bOk = TRUE;
	dwSize = sizeof(DWORD);

	bOk = HttpQueryInfo(hRequest, HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER, &dwStatus, &dwSize, 0);
	OutputDebugFormat(L"HTTP_STATUS %d (Ok:%d)\r\n", dwStatus, bOk);
}

