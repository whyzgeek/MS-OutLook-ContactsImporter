////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiSink.cpp
// Description: MAPI Advise Sink Wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include <MapiUtil.h>
#include "ExMapiSink.h"

////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CExMapiSink

CExMapiSink::CExMapiSink(LPNOTIFCALLBACK lpfnCallback,LPVOID lpvContext)
{
	m_lpfnCallback=lpfnCallback;
	m_lpvContext=lpvContext;
	m_nRef=0;
}

HRESULT CExMapiSink::QueryInterface(REFIID riid,LPVOID FAR* ppvObj)
{
	if(riid==IID_IUnknown) {
		*ppvObj=this;
		AddRef();
		return S_OK;
	}  
	return E_NOINTERFACE;
}

ULONG CExMapiSink::AddRef()
{
	return InterlockedIncrement(&m_nRef);
}

ULONG CExMapiSink::Release()
{
	ULONG ul=InterlockedDecrement(&m_nRef);
	if(!ul) delete this;
	return ul;
}

ULONG CExMapiSink::OnNotify(ULONG cNotification,LPNOTIFICATION lpNotifications)
{
	if(m_lpfnCallback) m_lpfnCallback(m_lpvContext,cNotification,lpNotifications);
	return 0;
}

