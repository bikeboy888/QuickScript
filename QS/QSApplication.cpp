// QSApplication.cpp : Implementation of CQSApplication

#include "stdafx.h"
#include "QSApplication.h"


// CQSApplication

STDMETHODIMP CQSApplication::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSApplication
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
