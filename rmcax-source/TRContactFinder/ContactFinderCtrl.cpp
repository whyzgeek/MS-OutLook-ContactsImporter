////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ContactFinderCtrl.cpp
// Description: Implementation of CContactFinderCtrl
//
// Developed by CallTrix Ltd, www.calltrix.com
// This is the main class
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ContactFinderCtrl.h"
#include "msi.h"
#include "nk2parser.h"
#include "ExMapi.h"
#include "CommonMethods.h"
#include "textfileio.h"

CContactFinderCtrl* __pTRCtrl = NULL;

CString ConvertContactListToJSON(CObArray& contacts)
{
	CString contactsJSON("{\"contacts\":[");
	for(int i=0;i<contacts.GetCount();i++)
	{
		CContact* pContact = (CContact*)contacts.GetAt(i);
		contactsJSON += pContact->ToJsonString();			
		if(i+1 <contacts.GetCount())
		{
			contactsJSON += ",";
		}
	}
	contactsJSON += "]}";
	return contactsJSON;
}


//
// Global Functions execute in seperate thread for fetching contacts
//
// For PAB



BOOL GetPABContacts(CString& contacts,CTRError& error)
{
	BOOL result = FALSE;

	CObArray contactArray;
	contacts = ConvertContactListToJSON(contactArray);	
	if(__pTRCtrl == NULL)
		return result;
	
	result = __pTRCtrl->m_pMapiEx->GetContacts(contactArray,error);
	__pTRCtrl->m_pMapiEx->Release();
		
	contacts = ConvertContactListToJSON(contactArray);	
	int count = contactArray.GetCount();
	for(int i=0;i<count;i++)
	{
		CContact* pContact = (CContact*)contactArray.GetAt(0);
		contactArray.RemoveAt(0);
		delete pContact;
	}
	
	return result;
}

UINT ThreadProcForGetPABContacts(LPVOID pParam)
{	
	__pTRCtrl = (CContactFinderCtrl*)pParam;
	CTRError* pEror = new CTRError();	
	CString contacts;
	if(GetPABContacts(contacts,*pEror))
	{		
		PostMessage(__pTRCtrl->m_hWnd,WM_THREAD1,(WPARAM)contacts.AllocSysString(),NULL);	
	}
	else
	{
		PostMessage(__pTRCtrl->m_hWnd,WM_ERROR,(WPARAM)pEror,NULL);	
	}

	return 0;
}

UINT ThreadProcForGetPABFolderCount(LPVOID pParam)
{	
	__pTRCtrl = (CContactFinderCtrl*)pParam;
	
	CString contacts;

	ULONG count = __pTRCtrl->m_pMapiEx->GetContactFolderCount();
	__pTRCtrl->m_pMapiEx->Release();

	if(count > 0)
	{		
		PostMessage(__pTRCtrl->m_hWnd,WM_THREAD2,(WPARAM)count,NULL);	
	}
	else
	{
		CTRError* pError = new CTRError();	
		pError->code = 107;
		pError->err = CString(L"TRContactFinder Cannot find private folders.");

		PostMessage(__pTRCtrl->m_hWnd,WM_ERROR,(WPARAM)pError,NULL);	
	}

	return 0;
}

UINT ThreadProcForGetPABContactCount(LPVOID pParam)
{	
	__pTRCtrl = (CContactFinderCtrl*)pParam;
	
	CString contacts;

	ULONG count = __pTRCtrl->m_pMapiEx->GetContactCount();	
	__pTRCtrl->m_pMapiEx->Release();
	if(count > 0)
	{		
		PostMessage(__pTRCtrl->m_hWnd,WM_THREAD3,(WPARAM)count,NULL);	
	}
	else
	{
		CTRError* pError = new CTRError();	
		pError->code = 107;
		pError->err = CString(L"TRContactFinder Cannot find private folders.");

		PostMessage(__pTRCtrl->m_hWnd,WM_ERROR,(WPARAM)pError,NULL);	
	}

	return 0;
}



//NK2
BOOL GetNK2Contacts(CString& contacts,CTRError& error)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CObArray contactsArray;
	contacts = ConvertContactListToJSON(contactsArray);	
	CString strFilterBy = __pTRCtrl->GetNK2SerachBy()  ;
	CString startWith = __pTRCtrl->GetNK2SerachStartWith();
	ULONG lPageSize;
	__pTRCtrl->get_NK2PageSize(&lPageSize);
	LONG lPageNo = __pTRCtrl->GetNK2PageNumber();

	CNK2Parser nk2Parser(__pTRCtrl);
	if(nk2Parser.GetNK2Contacts(strFilterBy,startWith,__pTRCtrl->FetchUnique(),lPageSize,lPageNo,contactsArray,error))
	{
		contacts = ConvertContactListToJSON(contactsArray);	
		int count = contactsArray.GetCount();
		for(int i=0;i<count;i++)
		{
			CContact* pContact = (CContact*)contactsArray.GetAt(0);
			contactsArray.RemoveAt(0);
			delete pContact;
		}
		return TRUE;
	}	
	return FALSE;
}

UINT ThreadProcForGetNK2Contacts(LPVOID pParam)
{
	__pTRCtrl = (CContactFinderCtrl*)pParam;
	CTRError* pEror = new CTRError();	
	CString contacts;
	if(GetNK2Contacts(contacts,*pEror))
	{	
		PostMessage(__pTRCtrl->m_hWnd,WM_THREAD4,(WPARAM)contacts.AllocSysString(),(LPARAM)NULL);			
	}
	else
	{
		PostMessage(__pTRCtrl->m_hWnd,WM_ERROR,(WPARAM)pEror,NULL);	
	}

	return 0;
}



UINT ThreadProcForGetNK2ContactsCount(LPVOID pParam)
{
	__pTRCtrl = (CContactFinderCtrl*)pParam;
	CTRError* pError = new CTRError();	

	CNK2Parser nk2Parser(__pTRCtrl);	
	int count = nk2Parser.GetNK2ContactsCount(*pError);	
	if(pError->code == 108)
	{
		PostMessage(__pTRCtrl->m_hWnd,WM_ERROR,(WPARAM)pError,NULL);	
	}
	else
	{	
		PostMessage(__pTRCtrl->m_hWnd,WM_THREAD5,(WPARAM)count,(LPARAM)NULL);			
	}
	
	return 0;
}


//<SiteLock>
//Uncomment following 3 lines for SiteLock
//const CContactFinderCtrl::SiteList CContactFinderCtrl::rgslTrustedSites[]=
//{	
//	{SiteList::Allow, L"http", L"www.google.com"}
//};



//
// ActiveX Control Methods and Properties
//

STDMETHODIMP CContactFinderCtrl::InitCtrl(ULONG* pInit)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	HRESULT hr = S_OK;
	*pInit = PreInit() ? 1 : 0;
	return hr;
}

STDMETHODIMP CContactFinderCtrl::StartPABContactSearch(BSTR bstrSearchBy, BSTR bstrStartWith, ULONG lPageNumber)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// TODO: Add your implementation code here	
	if(m_bMAPIInProgress)
	{
		CString error(L"TRContactFinder is in progress. To work with this function please abort the previous thread.");
		Fire_OnError(114,error.AllocSysString());						
		return S_OK;
	}

	if(PreInit())
	{		
		if(m_pMapiEx == NULL)
			m_pMapiEx = new CExMapi(this);
		if(!m_pMapiEx->IsLogin())
		{
			if(CExMapi::Init())
			{
				if(!m_pMapiEx->Login())
				{
					CString error(L"TRContactFinder Cannot find private folders.");
					Fire_OnError(107,error.AllocSysString());	
					return S_OK;
				}
			}
		}
		m_bMAPIInProgress = TRUE;
		m_bPABSearchStart = TRUE;
		m_bAbortPABSearch = FALSE;
		m_bstrPABSearchBy = bstrSearchBy;
		m_bstrPABSearchStartWith = bstrStartWith;
		m_lPABPageNumber = lPageNumber;
		AfxBeginThread(ThreadProcForGetPABContacts,(LPVOID)this);
	}
	return S_OK;
}
STDMETHODIMP CContactFinderCtrl::StartPABFolderCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here

	if(m_bMAPIInProgress)
	{
		CString error(L"TRContactFinder is in progress. To work with this function please abort running thread.");
		Fire_OnError(114,error.AllocSysString());						
		return S_OK;
	}
		
	if(PreInit())
	{
		if(m_pMapiEx == NULL)
			m_pMapiEx = new CExMapi(this);
		if(!m_pMapiEx->IsLogin())
		{
			if(CExMapi::Init())
			{
				if(!m_pMapiEx->Login())
				{
					CString error(L"TRContactFinder Cannot find private folders.");
					Fire_OnError(107,error.AllocSysString());						
					return S_OK;
				}
			}
		}		

		m_bMAPIInProgress = TRUE;	
		m_bPABFolderCountingStart = TRUE;
		AfxBeginThread(ThreadProcForGetPABFolderCount,(LPVOID)this);			
	}	
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::StartPABContactCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here

	if(m_bMAPIInProgress)
	{
		CString error(L"TRContactFinder is in progress. To work with this function please abort running thread.");
		Fire_OnError(114,error.AllocSysString());						
		return S_OK;
	}
	
	if(PreInit())
	{
		if(m_pMapiEx == NULL)
			m_pMapiEx = new CExMapi(this);
		if(!m_pMapiEx->IsLogin())
		{
			if(CExMapi::Init())
			{
				if(!m_pMapiEx->Login())
				{
					CString error(L"TRContactFinder Cannot find private folders.");
					Fire_OnError(107,error.AllocSysString());						
					return S_OK;
				}
			}
		}		

		m_bMAPIInProgress = TRUE;
		m_bPABContactCountingStart = TRUE;
		AfxBeginThread(ThreadProcForGetPABContactCount,(LPVOID)this);			
	}	

	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::StartNK2ContactSearch(BSTR bstrSearchBy, BSTR bstrStartWith, ULONG lPageNumber)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bNK2InProgress)
	{
		CString error(L"TRContactFinder NK2 Contact Reading in progress. To work with this function abort running thread.");
		Fire_OnError(115,error.AllocSysString());						
		return S_OK;
	}

	if(PreInit())
	{
		m_bNK2SearchStart = TRUE;
		m_bAbortNK2Search = FALSE;
		m_bstrNK2SearchBy = bstrSearchBy;
		m_bstrNK2SearchStartWith = bstrStartWith;
		m_lNK2PageNumber = lPageNumber;
		m_bNK2InProgress = TRUE;
		AfxBeginThread(ThreadProcForGetNK2Contacts,(LPVOID)this);
	}
	return S_OK;
}
STDMETHODIMP CContactFinderCtrl::StartNK2ContactCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bNK2InProgress)
	{
		CString error(L"TRContactFinder NK2 Contact Reading in progress. To work with this function abort running thread.");
		Fire_OnError(115,error.AllocSysString());						
		return S_OK;

	}
	if(PreInit())
	{
		m_bNK2ContactCountingStart = TRUE;
		m_bAbortNK2ContactCounting = FALSE;		
		m_bNK2InProgress = TRUE;
		AfxBeginThread(ThreadProcForGetNK2ContactsCount,(LPVOID)this);		
	}
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::Abort(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bPABSearchStart || m_bPABFolderCountingStart || m_bPABContactCountingStart)
	{
		m_bAbortPABSearch = m_bAbortPABFolderCounting = m_bAbortPABContactCounting = TRUE;
		m_bPABSearchStart = m_bPABContactCountingStart = m_bPABContactCountingStart = FALSE;

	}
	if(m_bNK2SearchStart || m_bNK2ContactCountingStart)
	{
		m_bAbortNK2Search = m_bAbortNK2ContactCounting = TRUE;
		m_bNK2SearchStart = m_bNK2ContactCountingStart = FALSE;
	}
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::AbortPABSearch(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here	
	if(m_bPABSearchStart)
	{
		m_bAbortPABSearch = TRUE;
		m_bPABSearchStart = FALSE;
	}
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::AbortPABFolderCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bPABFolderCountingStart)
	{
		m_bAbortPABFolderCounting = TRUE;
		m_bPABFolderCountingStart = FALSE;
	}

	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::AbortPABContactCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bPABContactCountingStart)
	{
		m_bAbortPABContactCounting = TRUE;
		m_bPABContactCountingStart = FALSE;
	}
	return S_OK;
}


STDMETHODIMP CContactFinderCtrl::AbortNK2Search(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bNK2SearchStart)
	{
		m_bAbortNK2Search = TRUE;
		m_bNK2SearchStart = FALSE;
	}
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::AbortNK2ContactCounting(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(m_bNK2ContactCountingStart)
	{
		m_bAbortNK2ContactCounting = TRUE;
		m_bNK2ContactCountingStart = FALSE;
	}
	return S_OK;
}
STDMETHODIMP CContactFinderCtrl::GetPABContactsCount(ULONG* pCount)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here	
	LONG count = 0;		
	/*if(PreInit())
	{
		if(m_pMapiEx == NULL)
			m_pMapiEx = new CExMapi(this);
		if(!m_pMapiEx->IsLogin())
		{
			if(CExMapi::Init())
			{
				if(!m_pMapiEx->Login())
				{
					CString error(L"TRContactFinder Cannot find private folders.");
					Fire_OnError(107,error.AllocSysString());	
					*pCount = 0;
					return S_OK;
				}
			}
		}
		
		count = m_pMapiEx->GetContactCount();				
	}*/
	*pCount = count;
	
	CString error(L"TRContactFinder does not support GetPABContactsCount method. Please call StartPABContactCounting method.");
	Fire_OnError(113,error.AllocSysString());						

	return S_OK;
}
STDMETHODIMP CContactFinderCtrl::GetPABFolderCount(ULONG* pCount)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	
	ULONG count = 0;	
	/*if(PreInit())
	{
		if(m_pMapiEx == NULL)
			m_pMapiEx = new CExMapi(this);
		if(!m_pMapiEx->IsLogin())
		{
			if(CExMapi::Init())
			{
				if(!m_pMapiEx->Login())
				{
					CString error(L"TRContactFinder Cannot find private folders.");
					Fire_OnError(107,error.AllocSysString());	
					*pCount = 0;
					return S_OK;
				}
			}
		}		
		count = m_pMapiEx->GetContactFolderCount();
		if(count == 0)
		{		
			CString err(L"TRContactFinder Cannot find private folders.");
			Fire_OnError(107,err.AllocSysString());	
		}		
	}*/

	*pCount = count;
	CString error(L"TRContactFinder does not support GetPABFolderCount method. Please call StartPABFolderCounting method.");
	Fire_OnError(112,error.AllocSysString());						

	return S_OK;
}
STDMETHODIMP CContactFinderCtrl::GetNK2ContactsCount(ULONG* pCount)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	/*if(PreInit())
	{
		CNK2Parser nk2Parser;	
		CTRError e;
		*pCount = nk2Parser.GetNK2ContactsCount(e);
		if(e.code == 108)
		{
			Fire_OnError(e.code,e.err.AllocSysString());
		}
	}*/
	*pCount = 0;
	CString error(L"TRContactFinder does not support GetNK2ContactsCount method. Please call StartNK2ContactCounting method.");
	Fire_OnError(116,error.AllocSysString());						

	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::get_FetchUnique(USHORT* pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	*pVal = (SHORT)m_bFetchUnique;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::put_FetchUnique(USHORT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	m_bFetchUnique = (BOOL)newVal;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::get_PABPageSize(ULONG* pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	*pVal = m_lPABPageSize;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::put_PABPageSize(ULONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	m_lPABPageSize = newVal;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::get_NK2PageSize(ULONG* pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	*pVal = m_lNK2PageSize;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::put_NK2PageSize(ULONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	m_lNK2PageSize = newVal;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::put_SubFolderLevel(ULONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(newVal == 0 || newVal > 4)
		newVal = 4;

	m_lSubFolderLevel = newVal;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::get_EnableLog(USHORT* pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	*pVal = m_bEnableLog ? 1 : 0;
	return S_OK;
}

STDMETHODIMP CContactFinderCtrl::put_EnableLog(USHORT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
	if(newVal == 0)
		m_bEnableLog = FALSE;
	else
		m_bEnableLog = TRUE;

	return S_OK;
}

//
// Class Methods
//
BOOL CContactFinderCtrl::PreInit()
{
	BOOL init = TRUE;
	CTRError err;
	
	if(IsOS(25) || IsOS(19) || IsOS(18) || IsOS(8))
	{			
		if(CCommonMethods::IsValidIE())
		{
			// TODO: Add your implementation code here
			if(!CCommonMethods::IsOutlookInstalled(err))
			{		
				init = FALSE;		
			}
			else if(!CCommonMethods::IsVaildOutlookVersion(err))
			{
				init = FALSE;

			}
			else if(!CCommonMethods::IsOutlookDefaultMailClient(err))
			{
				//MakeOutlookDefaultMailClient();
				init = FALSE;
			}			
		}
		else
		{
			err.code = 105;
			err.err = "TRContactFinder only support IE 6.* and above.";
			init = FALSE;
		}
	}
	else
	{
		err.code = 104;
		err.err = "TRContactFinder does not support current OS.";
		init = FALSE;
	}
	if(!init)
	{
		Fire_OnError(err.code,err.err.AllocSysString());	
	}
	return init;
}


LRESULT CContactFinderCtrl::OnFireFinishPABSearchContactsEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.		
	m_bAbortPABSearch = FALSE;
	m_bMAPIInProgress = FALSE;
	if(wParam != NULL)
	{				
		Fire_OnFinishPABContactsSearch((BSTR)wParam);
	}
	return TRUE;
}
LRESULT CContactFinderCtrl::OnFireFinishPABFolderCountEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.		
	m_bAbortPABFolderCounting = FALSE;
	m_bPABFolderCountingStart = FALSE;
	m_bMAPIInProgress = FALSE;	
	Fire_OnFinishPABFolderCounting((ULONG)wParam);		

	return TRUE;
}
LRESULT CContactFinderCtrl::OnFireFinishPABContactCountEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.		
	m_bAbortPABContactCounting = FALSE;
	m_bPABContactCountingStart = FALSE;
	m_bMAPIInProgress = FALSE;
	Fire_OnFinishPABContactCounting((ULONG)wParam);

	return TRUE;
}

LRESULT CContactFinderCtrl::OnFireFinishNK2SearchContactsEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.	
	m_bAbortNK2Search = FALSE;
	m_bNK2SearchStart = FALSE;
	m_bNK2InProgress = FALSE;
	if(wParam != NULL)
	{		
		Fire_OnFinishNK2ContactsSearch((BSTR)wParam);
	}
	return TRUE;
}

LRESULT CContactFinderCtrl::OnFireFinishNK2ContactsCountingEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.	
	m_bAbortNK2ContactCounting = FALSE;
	m_bNK2ContactCountingStart = FALSE;
	m_bNK2InProgress = FALSE;
	Fire_OnFinishNK2ContactCounting((ULONG)wParam);
	return TRUE;
}

LRESULT CContactFinderCtrl::OnFireErrorEvent(UINT uMsg,WPARAM wParam, LPARAM lParam,BOOL& bHandle)
{
	//This is a one-message-does-everything handler.  If wParam is not
	//NULL, that means our message has been sent to fire the event.	
	m_bAbortNK2Search = FALSE;
	if(wParam != NULL)
	{
		CTRError* pError = (CTRError*)(wParam);
		if(pError->code != 112 && pError->code != 113 && pError->code != 114)
		{
			m_bMAPIInProgress = FALSE;
		}
		if(pError->code != 115 && pError->code != 116)
		{
			m_bNK2InProgress = FALSE;
		}
		Fire_OnError(pError->code,pError->err.AllocSysString());
	}	
	return TRUE;
}


CString CContactFinderCtrl::GetPABSerachStartWith(void)
{
	return CString(m_bstrPABSearchStartWith);
}

CString CContactFinderCtrl::GetPABSerachBy(void)
{
	return CString(m_bstrPABSearchBy);
}

CString CContactFinderCtrl::GetNK2SerachStartWith(void)
{
	return CString(m_bstrNK2SearchStartWith);
}
CString CContactFinderCtrl::GetNK2SerachBy(void)
{
	return CString(m_bstrNK2SearchBy);
}

ULONG CContactFinderCtrl::GetPABPageNumber()
{
	return m_lPABPageNumber;
}
ULONG CContactFinderCtrl::GetNK2PageNumber()
{
	return m_lNK2PageNumber;
}
ULONG CContactFinderCtrl::GetSubFOlderLevel()
{
	return m_lSubFolderLevel;
}
BOOL CContactFinderCtrl::FetchUnique()
{
	return m_bFetchUnique;
}
BOOL CContactFinderCtrl::IsPABSearchAbort()
{
	return m_bAbortPABSearch;
}
BOOL CContactFinderCtrl::IsPABContactCoutningAbort()
{
	return m_bAbortPABContactCounting;
}
BOOL CContactFinderCtrl::IsPABFolderCoutningAbort()
{
	return m_bAbortPABFolderCounting;
}
BOOL CContactFinderCtrl::IsNK2SearchAbort()
{
	return m_bAbortNK2Search;
}




BOOL CContactFinderCtrl::IsNK2ContactCountingAbort()
{
	return m_bAbortNK2ContactCounting;
}

BOOL CContactFinderCtrl::IsEnableLog()
{
	return m_bEnableLog;
}


