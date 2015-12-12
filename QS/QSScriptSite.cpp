// QSScriptSite.cpp : Implementation of CQSScriptSite

#include "stdafx.h"
#include "QSScriptSite.h"


// CQSScriptSite

STDMETHODIMP CQSScriptSite::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSScriptSite
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
