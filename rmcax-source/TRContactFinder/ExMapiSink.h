#ifndef __EXMAPISINK_H__
#define __EXMAPISINK_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiSink.h
// Description: Extendned MAPI Advise Sink Wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

// CExMapiSink

class CExMapiSink : public IMAPIAdviseSink
{
public:
	CExMapiSink(LPNOTIFCALLBACK lpfnCallback,LPVOID lpvContext);

// Attributes
protected:
	LPNOTIFCALLBACK m_lpfnCallback;  
	LPVOID m_lpvContext;  
	LONG m_nRef;  

// IUnknown
public:
	STDMETHOD(QueryInterface)(REFIID riid,LPVOID FAR* ppvObj);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

// IMAPIAdviseSink
public:
	STDMETHOD_(ULONG, OnNotify)(ULONG cNotification,LPNOTIFICATION lpNotifications);
};  

#endif
