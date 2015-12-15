#pragma once

class CCoInit
{
public:
	CCoInit()
	{
		CoInitializeEx(NULL, COINIT_MULTITHREADED);
	}

	virtual ~CCoInit()
	{
		CoUninitialize();
	}
};
