////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ContactFinderCtrl.h
// Description: Declaration of the CContactFinderCtrl
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once
#include "resource.h"       // main symbols
#include <atlctl.h>
#include "TRContactFinder_i.h"
#include "_IContactFinderCtrlEvents_CP.h"
#include "ExMapi.h"
//#include "SiteLock.h"  //<SiteLock> Uncomment this line for SiteLock

//User Define Windows messages to fire events during thread execution
#define WM_THREAD1	WM_USER + 101 //Start PAB Contact Search Thread
#define WM_THREAD2	WM_USER + 102 //Start PAB Folder Counting Thread
#define WM_THREAD3	WM_USER + 103 //Start PAB Contact Counting Thread
#define WM_THREAD4	WM_USER + 104 //Start NK2 Contact Search Thread
#define WM_THREAD5	WM_USER + 105 //Start NK2 Contact Counting Thread
#define WM_ERROR	WM_USER + 106

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif


// CContactFinderCtrl
class ATL_NO_VTABLE CContactFinderCtrl :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IContactFinderCtrl, &IID_IContactFinderCtrl, &LIBID_TRContactFinderLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IPersistStreamInitImpl<CContactFinderCtrl>,
	public IOleControlImpl<CContactFinderCtrl>,
	public IOleObjectImpl<CContactFinderCtrl>,
	public IOleInPlaceActiveObjectImpl<CContactFinderCtrl>,
	public IViewObjectExImpl<CContactFinderCtrl>,
	public IOleInPlaceObjectWindowlessImpl<CContactFinderCtrl>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CContactFinderCtrl>,
	public CProxy_IContactFinderCtrlEvents<CContactFinderCtrl>,
	public IObjectWithSiteImpl<CContactFinderCtrl>,
	public IPersistStorageImpl<CContactFinderCtrl>,
	public ISpecifyPropertyPagesImpl<CContactFinderCtrl>,
	public IQuickActivateImpl<CContactFinderCtrl>,
#ifndef _WIN32_WCE
	public IDataObjectImpl<CContactFinderCtrl>,
#endif
	public IProvideClassInfo2Impl<&CLSID_ContactFinderCtrl, &__uuidof(_IContactFinderCtrlEvents), &LIBID_TRContactFinderLib>,	
	public IObjectSafetyImpl<CContactFinderCtrl, INTERFACESAFE_FOR_UNTRUSTED_CALLER>,

	//<SiteLock>
	//Un-Comment Following and Comment above line for SiteLock
	/*public IObjectSafetySiteLockImpl<CContactFinderCtrl, INTERFACESAFE_FOR_UNTRUSTED_CALLER>,*/

	public CComCoClass<CContactFinderCtrl, &CLSID_ContactFinderCtrl>,
	public CComControl<CContactFinderCtrl>
{
public:


	CContactFinderCtrl()
	{
	}

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_INSIDEOUT |
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST
)

DECLARE_REGISTRY_RESOURCEID(IDR_CONTACTFINDERCTRL)


BEGIN_COM_MAP(CContactFinderCtrl)
	COM_INTERFACE_ENTRY(IContactFinderCtrl)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObjectEx)
	COM_INTERFACE_ENTRY(IViewObject2)
	COM_INTERFACE_ENTRY(IViewObject)
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceObject)
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
	COM_INTERFACE_ENTRY(IOleControl)
	COM_INTERFACE_ENTRY(IOleObject)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
	COM_INTERFACE_ENTRY(IQuickActivate)
	COM_INTERFACE_ENTRY(IPersistStorage)
#ifndef _WIN32_WCE
	COM_INTERFACE_ENTRY(IDataObject)
#endif
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
	COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY_IID(IID_IObjectSafety, IObjectSafety)
	
	//<SiteLock>
	//Un-Comment following lines for SiteLock
	//COM_INTERFACE_ENTRY(IObjectSafety)
	//COM_INTERFACE_ENTRY(IObjectSafetySiteLock)

END_COM_MAP()

BEGIN_PROP_MAP(CContactFinderCtrl)
END_PROP_MAP()

BEGIN_CONNECTION_POINT_MAP(CContactFinderCtrl)
	CONNECTION_POINT_ENTRY(__uuidof(_IContactFinderCtrlEvents))
END_CONNECTION_POINT_MAP()

BEGIN_MSG_MAP(CContactFinderCtrl)
	MESSAGE_HANDLER(WM_THREAD1,OnFireFinishPABSearchContactsEvent)
	MESSAGE_HANDLER(WM_THREAD2,OnFireFinishPABFolderCountEvent)
	MESSAGE_HANDLER(WM_THREAD3,OnFireFinishPABContactCountEvent)
	MESSAGE_HANDLER(WM_THREAD4,OnFireFinishNK2SearchContactsEvent)
	MESSAGE_HANDLER(WM_THREAD5,OnFireFinishNK2ContactsCountingEvent)
	MESSAGE_HANDLER(WM_ERROR,OnFireErrorEvent)
	CHAIN_MSG_MAP(CComControl<CContactFinderCtrl>)
	DEFAULT_REFLECTION_HANDLER()
END_MSG_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
	{
		static const IID* arr[] =
		{
			&IID_IContactFinderCtrl,
		};

		for (int i=0; i<sizeof(arr)/sizeof(arr[0]); i++)
		{
			if (InlineIsEqualGUID(*arr[i], riid))
				return S_OK;
		}
		return S_FALSE;
	}

// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// IContactFinderCtrl
public:
	HRESULT OnDraw(ATL_DRAWINFO& di)
	{
		return S_OK;
	}

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{		
		m_bWindowOnly = true;
		m_bPABSearchStart = FALSE;
		m_bNK2SearchStart = FALSE;		
		m_bPABFolderCountingStart = FALSE;
		m_bPABContactCountingStart = FALSE;
		m_bAbortPABSearch = FALSE;
		m_bAbortPABFolderCounting = FALSE;
		m_bAbortPABContactCounting = FALSE;
		m_bAbortNK2Search = FALSE;		
		m_bNK2ContactCountingStart = FALSE;
		m_bAbortNK2ContactCounting = FALSE;		
		m_bFetchUnique = TRUE;
		m_bMAPIInProgress = FALSE;
		m_bNK2InProgress = FALSE;
		m_bEnableLog = FALSE;
		m_lPABContactCount = 0;
		m_lPABFolderCount = 0;
		m_lNK2ContactCount = 0;
		m_lPABPageSize = 50;
		m_lNK2PageSize = 50;
		m_lPABPageNumber = 1;
		m_lNK2PageNumber = 1;
		m_lSubFolderLevel = 4;

		m_pMapiEx = new CExMapi(this);
		
		return S_OK;
	}

	void FinalRelease()
	{		
		if(m_pMapiEx!=NULL)
		{			
			if(m_pMapiEx->IsLogin())
				m_pMapiEx->Logout();
			delete m_pMapiEx;
			CExMapi::Term();
		}
	}

//Private Methods and Properties
private:
	BSTR m_bstrPABContacts;
	BSTR m_bstrNK2Contacts;	
	BSTR m_bstrPABSearchStartWith;
	BSTR m_bstrPABSearchBy;
	BSTR m_bstrNK2SearchStartWith;
	BSTR m_bstrNK2SearchBy;

	BOOL m_bPABSearchStart;
	BOOL m_bAbortPABSearch;

	BOOL m_bPABFolderCountingStart;
	BOOL m_bAbortPABFolderCounting;

	BOOL m_bPABContactCountingStart;	
	BOOL m_bAbortPABContactCounting;

	BOOL m_bNK2SearchStart;	
	BOOL m_bAbortNK2Search;

	BOOL m_bNK2ContactCountingStart;
	BOOL m_bAbortNK2ContactCounting;

	BOOL m_bFetchUnique;
	
	BOOL m_bMAPIInProgress;
	BOOL m_bNK2InProgress;

	ULONG m_lPABContactCount;
	ULONG m_lPABFolderCount;
	ULONG m_lNK2ContactCount;
	ULONG m_lPABPageNumber;
	ULONG m_lNK2PageNumber;
	ULONG m_lPABPageSize;
	ULONG m_lNK2PageSize;	
	ULONG m_lSubFolderLevel;
	BOOL m_bEnableLog;

	LRESULT OnFireFinishPABSearchContactsEvent(UINT uMsg ,WPARAM wParam, LPARAM lParam,BOOL& bHandle);
	LRESULT OnFireFinishPABFolderCountEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle);
	LRESULT OnFireFinishPABContactCountEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle);
	LRESULT OnFireFinishNK2SearchContactsEvent(UINT uMsg ,WPARAM wParam, LPARAM lParam,BOOL& bHandle);
	LRESULT OnFireFinishNK2ContactsCountingEvent(UINT uMsg ,WPARAM wParam, LPARAM lParam,BOOL& bHandle);
	LRESULT OnFireErrorEvent(UINT uMsg ,WPARAM wParam, LPARAM lParam,BOOL& bHandle);

	BOOL PreInit();

//Public Methods and Properties exposed by this class for internal use
public:

	CExMapi* m_pMapiEx;		

	BOOL IsPABSearchAbort();
	BOOL IsPABFolderCoutningAbort();
	BOOL IsPABContactCoutningAbort();
	BOOL IsNK2SearchAbort();
	BOOL IsNK2ContactCountingAbort();
	BOOL FetchUnique();

	CString GetPABSerachStartWith(void);
	CString GetPABSerachBy(void);
	CString GetNK2SerachStartWith(void);
	CString GetNK2SerachBy(void);
	
	ULONG GetPABPageNumber();
	ULONG GetNK2PageNumber();
	ULONG GetSubFOlderLevel();
	BOOL IsEnableLog();

	//<SiteLock>
	//Un-Comment following lines for SiteLock
	//static const SiteList rgslTrustedSites[2];
	//static const DWORD dwControlLifespan = 10;
	
//Public Methods and Properties exposed by this activeX control
public:
	//Methods
	STDMETHOD(InitCtrl)(ULONG* pInit);
	STDMETHOD(StartPABFolderCounting)(void);
	STDMETHOD(StartPABContactCounting)(void);
	STDMETHOD(StartPABContactSearch)(BSTR bstrSearchBy, BSTR bstrStartWith, ULONG lPageNumber);
	STDMETHOD(StartNK2ContactCounting)(void);
	STDMETHOD(StartNK2ContactSearch)(BSTR bstrSearchBy, BSTR bstrStartWith, ULONG lPageNumber);	
	STDMETHOD(Abort)(void);
	STDMETHOD(AbortPABSearch)(void);
	STDMETHOD(AbortPABFolderCounting)(void);
	STDMETHOD(AbortPABContactCounting)(void);
	STDMETHOD(AbortNK2Search)(void);	
	STDMETHOD(AbortNK2ContactCounting)(void);

	//Obsolete Methods
	STDMETHOD(GetPABContactsCount)(ULONG* pCount);
	STDMETHOD(GetPABFolderCount)(ULONG* pCount);
	STDMETHOD(GetNK2ContactsCount)(ULONG* pCount);

	//Properties
	STDMETHOD(get_FetchUnique)(USHORT* pVal);
	STDMETHOD(put_FetchUnique)(USHORT newVal);
	STDMETHOD(get_PABPageSize)(ULONG* pVal);
	STDMETHOD(put_PABPageSize)(ULONG newVal);
	STDMETHOD(get_NK2PageSize)(ULONG* pVal);
	STDMETHOD(put_NK2PageSize)(ULONG newVal);	
	STDMETHOD(put_SubFolderLevel)(ULONG newVal);
	STDMETHOD(get_EnableLog)(USHORT* pVal);
	STDMETHOD(put_EnableLog)(USHORT newVal);
};

OBJECT_ENTRY_AUTO(__uuidof(ContactFinderCtrl), CContactFinderCtrl)
