// QSNet.cpp : Implementation of CQSNet

#include "stdafx.h"
#include "QSNet.h"


// CQSNet

STDMETHODIMP CQSNet::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSNet
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
