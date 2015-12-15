// QSFile.cpp : Implementation of CQSFile

#include "stdafx.h"
#include "QSFile.h"


// CQSFile

STDMETHODIMP CQSFile::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSFile
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
