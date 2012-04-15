////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapi.cpp
// Description: Windows Extended MAPI class 
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include <MapiUtil.h>
#include <comdef.h>
#include "ExMapi.h"
#include "ExMapiSink.h"
#include "ExMapiNamedProperty.h"
#include "ContactFinderCtrl.h"
#include "Contact.h"

#pragma comment (lib,"mapi32.lib")
#pragma comment(lib,"Ole32.lib")

#define PR_EMAIL1_ADDRESS PROP_TAG( PT_TSTRING,	0x806A)
#define PR_EMAIL2_ADDRESS PROP_TAG( PT_TSTRING,	0x806B)
#define PR_EMAIL3_ADDRESS PROP_TAG( PT_TSTRING,	0x806C)

/////////////////////////////////////////////////////////////
// CExMapi

#ifdef UNICODE
int CExMapi::cm_nMAPICode=MAPI_UNICODE;
#else
int CExMapi::cm_nMAPICode=0;
#endif

CExMapi::CExMapi(CContactFinderCtrl* pCtrl)
{	
	m_pSession=NULL;
	m_pMsgStore=NULL;
	m_pFolder=NULL;	
	m_pContents=NULL;
	m_pHierarchy=NULL;
	m_sink=0;	

	m_pCtrl = pCtrl;		
}

CExMapi::~CExMapi()
{
	Logout();	
}

BOOL CExMapi::Init(BOOL bMultiThreadedNotifcations)
{			
#ifdef _WIN32_WCE
	if(CoInitializeEx(NULL, COINIT_MULTITHREADED)!=S_OK) return FALSE;
#endif
	if(bMultiThreadedNotifcations) {
		MAPIINIT_0 MAPIInit={ 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		if(MAPIInitialize(&MAPIInit)!=S_OK) return FALSE;
	} else {
		if(MAPIInitialize(NULL)!=S_OK) return FALSE;
	}
	return TRUE;
}

void CExMapi::Term()
{
	MAPIUninitialize();
#ifdef _WIN32_WCE
	CoUninitialize();
#endif

}

BOOL CExMapi::Login(LPCTSTR szProfileName)
{
	HRESULT hr = MAPILogonEx(
		NULL,
		(LPTSTR)szProfileName,
		NULL,
		MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_USE_DEFAULT | MAPI_UNICODE,
		&m_pSession
		);

	return (hr==S_OK);
}

BOOL CExMapi::IsLogin()
{
	return m_pSession != NULL;
}

void CExMapi::Logout()
{
	Release();
	RELEASE(m_pSession);		
}

void CExMapi::Release()
{
	if(m_sink) {
		if(m_pMsgStore) m_pMsgStore->Unadvise(m_sink);
		m_sink=0;
	}

	RELEASE(m_pHierarchy);
	RELEASE(m_pContents);
	RELEASE(m_pFolder);
	RELEASE(m_pMsgStore);
}

//call with ulEventMask set to ALL notifications ORed together, only one Advise Sink is used.
BOOL CExMapi::Notify(LPNOTIFCALLBACK lpfnCallback,LPVOID lpvContext,ULONG ulEventMask)
{
	if(GetMessageStoreSupport()&STORE_NOTIFY_OK) {
		if(m_sink) m_pMsgStore->Unadvise(m_sink);
		CExMapiSink* pAdviseSink=new CExMapiSink(lpfnCallback,lpvContext);
		if(m_pMsgStore->Advise(0,NULL,ulEventMask,pAdviseSink,&m_sink)==S_OK) return TRUE;
		delete pAdviseSink;
		m_sink=0;
	}
	return FALSE;
}

// if I try to use MAPI_UNICODE when UNICODE is defined I get the MAPI_E_BAD_CHARWIDTH 
// error so I force narrow strings here
LPCTSTR CExMapi::GetProfileName()
{
	if(!m_pSession) return NULL;

	static CString strProfileName;
	LPSRowSet pRows=NULL;
	const int nProperties=2;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME_A, PR_RESOURCE_TYPE}};

	IMAPITable*	pStatusTable;
	if(m_pSession->GetStatusTable(0,&pStatusTable)==S_OK) {
		if(pStatusTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) {
			while(TRUE) {
				if(pStatusTable->QueryRows(1,0,&pRows)!=S_OK) MAPIFreeBuffer(pRows);
				else if(pRows->cRows!=1) FreeProws(pRows);
				else if(pRows->aRow[0].lpProps[1].Value.ul==MAPI_SUBSYSTEM) {
					strProfileName=(LPSTR)GetValidString(pRows->aRow[0].lpProps[0]);
					FreeProws(pRows);
				} else {
					FreeProws(pRows);
					continue;
				}
				break;
			}
		}
		RELEASE(pStatusTable);
	}
	return strProfileName;
}

BOOL CExMapi::OpenDefaultMessageStore(LPCTSTR szStore,ULONG ulFlags)
{
	if(!m_pSession) return FALSE;

	m_ulMDBFlags=ulFlags;

	LPSRowSet pRows = NULL;
	const int nProperties = 3;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME, PR_ENTRYID, PR_DEFAULT_STORE}};

	BOOL bResult = FALSE;
	IMAPITable*	pMsgStoresTable;
	if(m_pSession->GetMsgStoresTable(0, &pMsgStoresTable) == S_OK) 
	{
		if(pMsgStoresTable->SetColumns((LPSPropTagArray)&Columns, 0) == S_OK) 
		{
			while(TRUE) 
			{
				if(pMsgStoresTable->QueryRows(1,0,&pRows) != S_OK)
					MAPIFreeBuffer(pRows);
				else if(pRows->cRows!=1) 
					FreeProws(pRows);
				else 
				{
					if(!szStore) 
					{ 
						if(pRows->aRow[0].lpProps[2].Value.b) bResult=TRUE;
					} 
					else 
					{
						CString strStore = GetValidString(pRows->aRow[0].lpProps[0]);
						if(strStore.Find(szStore) != -1) 
							bResult = TRUE;
					}
					if(!bResult) 
					{
						FreeProws(pRows);
						continue;
					}
				}
				break;
			}
			if(bResult) 
			{
				RELEASE(m_pMsgStore);
				bResult = (m_pSession->OpenMsgStore(NULL,pRows->aRow[0].lpProps[1].Value.bin.cb,(ENTRYID*)pRows->aRow[0].lpProps[1].Value.bin.lpb,NULL,MDB_NO_DIALOG | MAPI_BEST_ACCESS,&m_pMsgStore) == S_OK);
				FreeProws(pRows);
			}
		}
		RELEASE(pMsgStoresTable);
	}
	return bResult;
}


ULONG CExMapi::GetMessageStoreSupport()
{
	if(!m_pMsgStore) return FALSE;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	ULONG rgTags[]={ 1, PR_STORE_SUPPORT_MASK };
	ULONG ulSupport=0;

	if(m_pMsgStore->GetProps((LPSPropTagArray) rgTags, CExMapi::cm_nMAPICode, &cValues, &props)==S_OK) {
		ulSupport=props->Value.ul;
		MAPIFreeBuffer(props);
	}
	return ulSupport;
}

LPMAPIFOLDER CExMapi::OpenFolder(unsigned long ulFolderID,BOOL bInternal)
{
	if(!m_pMsgStore) return NULL;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	DWORD dwObjType;
	ULONG rgTags[]={ 1, ulFolderID };
	LPMAPIFOLDER pFolder;

	if(m_pMsgStore->GetProps((LPSPropTagArray) rgTags, cm_nMAPICode, &cValues, &props)!=S_OK) return NULL;
	m_pMsgStore->OpenEntry(props[0].Value.bin.cb,(LPENTRYID)props[0].Value.bin.lpb, NULL, m_ulMDBFlags, &dwObjType,(LPUNKNOWN*)&pFolder);
	MAPIFreeBuffer(props);

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CExMapi::OpenSpecialFolder(unsigned long ulFolderID,BOOL bInternal)
{
	LPMAPIFOLDER pInbox=OpenInbox(FALSE);
	if(!pInbox || !m_pMsgStore) return FALSE;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	DWORD dwObjType;
	ULONG rgTags[]={ 1, ulFolderID };
	LPMAPIFOLDER pFolder;

	if(pInbox->GetProps((LPSPropTagArray) rgTags, cm_nMAPICode, &cValues, &props)!=S_OK) return NULL;
	m_pMsgStore->OpenEntry(props[0].Value.bin.cb,(LPENTRYID)props[0].Value.bin.lpb, NULL, m_ulMDBFlags, &dwObjType,(LPUNKNOWN*)&pFolder);
	MAPIFreeBuffer(props);
	RELEASE(pInbox);

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CExMapi::OpenRootFolder(BOOL bInternal)
{
	return OpenFolder(PR_IPM_SUBTREE_ENTRYID,bInternal);
}

LPMAPIFOLDER CExMapi::OpenInbox(BOOL bInternal)
{
	if(!m_pMsgStore) return NULL;

#ifdef _WIN32_WCE
	return OpenFolder(PR_CE_IPM_INBOX_ENTRYID);
#else
	ULONG cbEntryID;
	LPENTRYID pEntryID;
	DWORD dwObjType;
	LPMAPIFOLDER pFolder;

	if(m_pMsgStore->GetReceiveFolder(NULL,0,&cbEntryID,&pEntryID,NULL)!=S_OK) return NULL;
	m_pMsgStore->OpenEntry(cbEntryID,pEntryID, NULL, m_ulMDBFlags,&dwObjType,(LPUNKNOWN*)&pFolder);
	MAPIFreeBuffer(pEntryID);
#endif

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CExMapi::OpenContacts(BOOL bInternal)
{
	return OpenSpecialFolder(PR_IPM_CONTACT_ENTRYID,bInternal);
}


LPMAPITABLE CExMapi::GetHierarchy(LPMAPIFOLDER pFolder)
{
	if(!pFolder) {
		pFolder=m_pFolder;
		if(!pFolder) return NULL;
	}
	RELEASE(m_pHierarchy);
	if(pFolder->GetHierarchyTable(0,&m_pHierarchy)!=S_OK) return NULL;

	const int nProperties=2;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME, PR_ENTRYID}};
	if(m_pHierarchy->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) return m_pHierarchy;
	return NULL;
}

LPMAPIFOLDER CExMapi::GetNextSubFolder(CString& strFolderName,LPMAPIFOLDER pFolder)
{
	if(!m_pHierarchy) return NULL;
	if(!pFolder) {
		pFolder=m_pFolder;
		if(!pFolder) return FALSE;
	}

	DWORD dwObjType;
	LPSRowSet pRows=NULL;	

	LPMAPIFOLDER pSubFolder=NULL;
	if(m_pHierarchy->QueryRows(1,0,&pRows)==S_OK) 
	{
		if(pRows->cRows) 
		{
			if(pFolder->OpenEntry(pRows->aRow[0].lpProps[PROP_ENTRYID].Value.bin.cb,(LPENTRYID)pRows->aRow[0].lpProps[PROP_ENTRYID].Value.bin.lpb, NULL, MAPI_MODIFY, &dwObjType,(LPUNKNOWN*)&pSubFolder)==S_OK) 
			{				
				strFolderName=CString(GetValidString(pRows->aRow[0].lpProps[0]));												
			}
		}
		FreeProws(pRows);
		MAPIFreeBuffer(pRows);
	}		
	return pSubFolder;
}

// High Level function to open a sub folder by iterating recursively (DFS) over all folders 
// (use instead of manually calling GetHierarchy and GetNextSubFolder)
LPMAPIFOLDER CExMapi::OpenSubFolder(LPCTSTR szSubFolder,LPMAPIFOLDER pFolder)
{
	LPMAPIFOLDER pSubFolder=NULL;
	LPMAPITABLE pHierarchy;

	RELEASE(m_pHierarchy);
	pHierarchy=GetHierarchy(pFolder);
	if(pHierarchy) 
	{
		CString strFolder;
		LPMAPIFOLDER pRecurse=NULL;
		do 
		{
			RELEASE(pSubFolder);
			m_pHierarchy=pHierarchy;
			pSubFolder = GetNextSubFolder(strFolder,pFolder);
			if(pSubFolder) 
			{
				if(!strFolder.CompareNoCase(szSubFolder)) 
					break;

				m_pHierarchy = NULL; // so we don't release it in subsequent drilldown
				pRecurse = OpenSubFolder(szSubFolder,pSubFolder);
				if(pRecurse) 
				{
					RELEASE(pSubFolder);
					pSubFolder = pRecurse;
					break;
				}
			}
		} while(pSubFolder);
		RELEASE(pHierarchy);
		m_pHierarchy=NULL;
	}
	// this may occur many times depending on how deep the recursion is; make sure we haven't already assigned m_pFolder
	if(pSubFolder && m_pFolder!=pSubFolder) {
		RELEASE(m_pFolder);
		m_pFolder=pSubFolder;
	}
	return pSubFolder;
} 

BOOL CExMapi::GetContents(LPMAPIFOLDER pFolder)
{
	if(!pFolder) {
		pFolder=m_pFolder;
		if(!pFolder) return FALSE;
	}
	RELEASE(m_pContents);
	if(pFolder->GetContentsTable(CExMapi::cm_nMAPICode,&m_pContents)!=S_OK) return FALSE;

	const int nProperties=MESSAGE_COLS;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_MESSAGE_FLAGS, PR_ENTRYID }};
	return (m_pContents->SetColumns((LPSPropTagArray)&Columns,0)==S_OK);
}

int CExMapi::GetRowCount()
{
	ULONG ulCount;
	if(!m_pContents || m_pContents->GetRowCount(0,&ulCount)!=S_OK) return -1;
	return ulCount;
}

BOOL CExMapi::RestrictContents(LPCTSTR pszSearchBy,LPCTSTR pszStartWith)
{
	if(m_pContents == NULL) 
		return FALSE;
	
	HRESULT hr = S_OK;	    	   
    	
    //if(strcmpi(pszFilter,"")!=0)
    if(wcscmp(pszStartWith,L"")!=0)
    {		
		SRestriction sres,srlevel1[2];
		SPropValue spvDisplay;

		//Create Restriction
		sres.rt = RES_AND;
		sres.res.resAnd.cRes = 1;
		sres.res.resAnd.lpRes = srlevel1;

		ULONG propTag = PR_DISPLAY_NAME;
		if(wcscmp(pszSearchBy,L"Email")==0)
		{
			propTag = PR_EMAIL1_ADDRESS;			
		}
		spvDisplay.ulPropTag = propTag;
		spvDisplay.Value.lpszW = (wchar_t *)pszStartWith;
		
		srlevel1[0].rt = RES_CONTENT;
		srlevel1[0].res.resContent.ulFuzzyLevel = FL_PREFIX | FL_IGNORECASE;
		srlevel1[0].res.resContent.ulPropTag = propTag;
		srlevel1[0].res.resContent.lpProp = &spvDisplay;

		hr = m_pContents->Restrict(&sres,0);
    }
    else
    {
 	    hr = m_pContents->Restrict(NULL,0);
    }
	return hr==S_OK;
}

BOOL CExMapi::SortContents(ULONG ulSortParam,ULONG ulSortField)
{
	if(!m_pContents) return FALSE;

	SizedSSortOrderSet(1, SortColums) = {1, 0, 0, {{ulSortField,ulSortParam}}};
	return (m_pContents->SortTable((LPSSortOrderSet)&SortColums,0)==S_OK);
}

BOOL CExMapi::GetNextContact(CExMapiContact& contact)
{
	if(!m_pContents) return FALSE;

	LPSRowSet pRows=NULL;
	BOOL bResult=FALSE;
	while(m_pContents->QueryRows(1,0,&pRows)==S_OK) {
		if(pRows->cRows) bResult=contact.Open(this,pRows->aRow[0].lpProps[PROP_ENTRYID].Value.bin);
		FreeProws(pRows);
		break;
	}
	MAPIFreeBuffer(pRows);
	return bResult;
}



// sometimes the string in prop is invalid, causing unexpected crashes
LPCTSTR CExMapi::GetValidString(SPropValue& prop)
{
	LPCTSTR s=prop.Value.LPSZ;
	if(s && !::IsBadStringPtr(s,(UINT_PTR)-1)) return s;
	return NULL;
}

// special case of GetValidString to take the narrow string in UNICODE
void CExMapi::GetNarrowString(SPropValue& prop,CString& strNarrow)
{
	LPCTSTR s=GetValidString(prop);
	if(!s) strNarrow=_T("");
	else {
#ifdef UNICODE
		// VS2005 can copy directly
		if(_MSC_VER>=1400) {
			strNarrow=(char*)s;
		} else {
			WCHAR wszWide[256];
			MultiByteToWideChar(CP_ACP,0,(LPCSTR)s,-1,wszWide,255);
			strNarrow=wszWide;
		}
#else
		strNarrow=s;
#endif
	}
}

BOOL CExMapi::IsContactFolder(LPMAPIFOLDER lpFolder)
{
	if(!lpFolder) return FALSE;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	ULONG rgTags[]={ 1, PR_CONTAINER_CLASS };

	if(SUCCEEDED(lpFolder->GetProps((LPSPropTagArray) rgTags, CExMapi::cm_nMAPICode, &cValues, &props)))
	{
		CString containerClass = CExMapi::GetValidString(*props);
		if(containerClass == L"IPF.Contact")
			return TRUE;
	}
	return FALSE;
}

CString CExMapi::GetFolderName(LPMAPIFOLDER lpFolder)
{
	CString folderName = L"";
	if(!lpFolder) return folderName;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	ULONG rgTags[]={ 1, PR_DISPLAY_NAME };

	if(SUCCEEDED(lpFolder->GetProps((LPSPropTagArray) rgTags, CExMapi::cm_nMAPICode, &cValues, &props)))
	{
		folderName = CExMapi::GetValidString(*props);		
	}
	return folderName;
}

LPMAPITABLE CExMapi::GetMessageStoresTable()
{
	if(!m_pSession) return FALSE;
	LPMAPITABLE lpMsgStoresTable = NULL;
	RELEASE(m_pMsgStore);
	if(SUCCEEDED(m_pSession->GetMsgStoresTable(0, &lpMsgStoresTable)))
		return lpMsgStoresTable;
	return NULL;
}

BOOL CExMapi::OpenNextMessageStore(LPMAPITABLE lpMsgStoreTable, ULONG ulFlags)
{
	if(!m_pSession) return FALSE;
	if(!lpMsgStoreTable) return FALSE;

	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogPAB(L"Function OpenNextMessageStore START.");
	}

	
	m_ulMDBFlags = ulFlags;

	LPSRowSet pRows = NULL;
	const int nProperties = 1;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_ENTRYID}};

	BOOL bResult = FALSE;	
	try
	{
		if(lpMsgStoreTable->SetColumns((LPSPropTagArray)&Columns, 0) == S_OK) 
		{
			while(TRUE) 
			{
				bResult = FALSE;
				if(lpMsgStoreTable->QueryRows(1,0,&pRows) != S_OK)
					MAPIFreeBuffer(pRows);
				else if(pRows->cRows!=1) 
					FreeProws(pRows);
				else 
				{					
					LPMDB pMsgStore = NULL;

					ULONG cbEntryId = pRows->aRow[0].lpProps[0].Value.bin.cb;
					LPENTRYID lpEntryId = (ENTRYID*)pRows->aRow[0].lpProps[0].Value.bin.lpb;
					
					bResult = (m_pSession->OpenMsgStore(NULL,
						cbEntryId,
						lpEntryId,
						NULL,
						MDB_NO_DIALOG | MAPI_BEST_ACCESS,
						&pMsgStore) == S_OK);				

					if(bResult)
					{
						if(m_pMsgStore)
						{
							//check if new open msg store is the same as opened msg store.
							if(AreSame(m_pMsgStore,pMsgStore))
							{
								if(doLog)
								{
									m_logHelper.LogPAB(L"Message Store is already opend.");
								}

								RELEASE(pMsgStore);
								FreeProws(pRows);							
								continue;
							}
							else
							{
								RELEASE(m_pMsgStore);
								m_pMsgStore = pMsgStore;				
								bResult = TRUE;
							}										
						}
						else
						{						
							m_pMsgStore = pMsgStore;				
						}				
					}				
					FreeProws(pRows);
				}
				break;
			}
		}		
	}
	catch(_com_error& e)
	{
		if(doLog)
		{
			m_logHelper.LogPAB(L"OpenNextMessageStore Exception:");
			m_logHelper.LogPAB(e.ErrorMessage());
		}
		return FALSE;
	}
	if(doLog)
	{
		m_logHelper.LogPAB(L"Function OpenNextMessageStore END.");
	}

	return bResult;
}
BOOL CExMapi::AreSame(LPMAPIPROP lpMapiProp1,LPMAPIPROP lpMapiProp2)
{
	if(!m_pSession) return FALSE;
	BOOL bResult=FALSE;

	if(lpMapiProp1 != NULL && lpMapiProp2 != NULL)
	{
		ULONG cbEntryId1;
		ULONG cbEntryId2;
		LPENTRYID lpEntryId1 = NULL;	
		LPENTRYID lpEntryId2 = NULL;

		SPropValue * rgprops1 = NULL;
		SPropValue * rgprops2 = NULL;
		ULONG rgTags[] = {1, PR_ENTRYID};
		ULONG cCount1   = 0;
		ULONG cCount2   = 0;
		// Get the message store entryid properties.
		if(SUCCEEDED(lpMapiProp1->GetProps((LPSPropTagArray) rgTags, MAPI_UNICODE, &cCount1, &rgprops1)))
		{
			cbEntryId1 = rgprops1[0].Value.bin.cb;
			lpEntryId1 = (ENTRYID*)rgprops1[0].Value.bin.lpb;
		}

		if(SUCCEEDED(lpMapiProp2->GetProps((LPSPropTagArray) rgTags, MAPI_UNICODE, &cCount2, &rgprops2)))
		{
			cbEntryId2 = rgprops2[0].Value.bin.cb;
			lpEntryId2 = (ENTRYID*)rgprops2[0].Value.bin.lpb;
		}

		if(lpEntryId1 != NULL && lpEntryId2 != NULL)
		{
			ULONG lResult;
			if(SUCCEEDED(m_pSession->CompareEntryIDs(cbEntryId1,lpEntryId1,cbEntryId2,lpEntryId2,0,&lResult)))
			{
				bResult = (BOOL)lResult;
			}			
		}

		MAPIFreeBuffer(rgprops1);
		MAPIFreeBuffer(rgprops2);
	}
	return bResult;
}
ULONG __recursionCount = 0;
LPMAPIFOLDER CExMapi::GetContactFolderCount(LPMAPIFOLDER pFolder,ULONG& count,BOOL doLog)
{
	
	__recursionCount++;
	wchar_t buffer[5];
	_itow(__recursionCount,buffer,10);
	if(doLog)
	{
		if(__recursionCount > 1)
			m_logHelper.LogPAB(L"-------------------------------------");
		m_logHelper.LogPAB(L"Recursive Call No: "+CString(buffer));
		m_logHelper.LogPAB(L"-------------------------------------");
	}	
	LPMAPIFOLDER pSubFolder=NULL;
	LPMAPITABLE pHierarchy;
	CString strFolder(L"");
		
	try
	{
		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 1 START");
		
		RELEASE(m_pHierarchy);		

		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 1 END");
		
		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder) START");

		pHierarchy = GetHierarchy(pFolder);
		
		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder) END");
		if(pHierarchy) 
		{
			LPMAPIFOLDER pRecurse = NULL;
			do 
			{
				if(doLog)
					m_logHelper.LogPAB(L"RELEASE(pSubFolder) 1 START");
								
				RELEASE(pSubFolder);
				
				if(doLog)
					m_logHelper.LogPAB(L"RELEASE(pSubFolder) 1 END");

				m_pHierarchy = pHierarchy;				

				if(doLog)
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder) START");
				
				pSubFolder = GetNextSubFolder(strFolder,pFolder);
				//AfxMessageBox(strFolder);

				if(doLog)
				{
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder) END");
					m_logHelper.LogPAB(L"FolderName: "+strFolder);
				}
				
				if(pSubFolder) 
				{
					if(strFolder == L"Deleted Items")
						continue;

					if(doLog)
						m_logHelper.LogPAB(L"if(IsContactFolder(pSubFolder)) START");

					if(IsContactFolder(pSubFolder))
					{
						count++;
					}

					if(doLog)
						m_logHelper.LogPAB(L"if(IsContactFolder(pSubFolder)) END");

					m_pHierarchy = NULL; // so we don't release it in subsequent drilldown
										
					if(__recursionCount < m_pCtrl->GetSubFOlderLevel())
					{
						if(doLog)
							m_logHelper.LogPAB(L"pRecurse = GetContactFolderCount(pSubFolder,count) START");

						pRecurse = GetContactFolderCount(pSubFolder,count,doLog);

						if(doLog)
							m_logHelper.LogPAB(L"pRecurse = GetContactFolderCount(pSubFolder,count) END");
					}


					if(pRecurse) 
					{
						if(doLog)
							m_logHelper.LogPAB(L"RELEASE(pSubFolder) 2 START");

						RELEASE(pSubFolder);

						if(doLog)
							m_logHelper.LogPAB(L"RELEASE(pSubFolder) 2 END");

						pSubFolder = pRecurse;
						break;
					}
				}
				else
				{
					__recursionCount = 1;
				}
			} while(pSubFolder && !m_pCtrl->IsPABFolderCoutningAbort());
			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 2 START");

			RELEASE(pHierarchy);

			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 2 END");

			m_pHierarchy=NULL;
		}
		// this may occur many times depending on how deep the recursion is; make sure we haven't already assigned m_pFolder
		if(pSubFolder && m_pFolder!=pSubFolder) 
		{
			if(doLog)			
				m_logHelper.LogPAB(L"RELEASE(m_pFolder) START");

			RELEASE(m_pFolder);

			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pFolder) END");
			m_pFolder=pSubFolder;
		}
	}
	catch(_com_error &e)
	{
		if(doLog)
		{
			m_logHelper.LogPAB(L"GetContactFolderCount Exception Catch START:");
			m_logHelper.LogPAB(L"FolderName: " + strFolder );
			m_logHelper.LogPAB(e.ErrorMessage());
			m_logHelper.LogPAB(L"GetContactFolderCount Exception Catch END:");
		}
	}
	return pSubFolder;
} 


LPMAPIFOLDER CExMapi::GetContactCount(LPMAPIFOLDER pFolder,ULONG& count, BOOL doLog)
{
	__recursionCount++;
	wchar_t buffer[5];
	_itow(__recursionCount,buffer,10);
	if(doLog)
	{
		if(__recursionCount > 1)
			m_logHelper.LogPAB(L"-------------------------------------");
		m_logHelper.LogPAB(L"Recursive Call No: "+CString(buffer));
		m_logHelper.LogPAB(L"-------------------------------------");
	}
	CString strFolder;

	
	LPMAPIFOLDER pSubFolder=NULL;
	LPMAPITABLE pHierarchy;
	try
	{
		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 1 START");

		RELEASE(m_pHierarchy);

		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 1 END");

		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder) START");

		pHierarchy = GetHierarchy(pFolder);

		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder) END");

		if(pHierarchy) 
		{
			LPMAPIFOLDER pRecurse = NULL;
			do 
			{
				if(doLog)
					m_logHelper.LogPAB(L"RELEASE(pSubFolder) 1 START");

				RELEASE(pSubFolder);

				if(doLog)
					m_logHelper.LogPAB(L"RELEASE(pSubFolder) 1 END");

				m_pHierarchy = pHierarchy;
				if(doLog)
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder) START");

				pSubFolder = GetNextSubFolder(strFolder,pFolder);

				if(doLog)
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder) END");
				
				if(pSubFolder) 
				{
					if(strFolder == L"Deleted Items")
						continue;
					if(doLog)
						m_logHelper.LogPAB(L"if(IsContactFolder(pSubFolder)) START");

					if(IsContactFolder(pSubFolder))
					{
						if(doLog)
							m_logHelper.LogPAB(L"if(GetContents(pSubFolder)) START");
						if(GetContents(pSubFolder))
						{
							if(doLog)
								m_logHelper.LogPAB(L"GetRowCount() START");

							count += GetRowCount();

							if(doLog)
								m_logHelper.LogPAB(L"GetRowCount() END");
						}
						if(doLog)
							m_logHelper.LogPAB(L"if(GetContents(pSubFolder)) END");
					}
					if(doLog)
						m_logHelper.LogPAB(L"if(IsContactFolder(pSubFolder)) END");

					m_pHierarchy = NULL; // so we don't release it in subsequent drilldown

					if(__recursionCount < m_pCtrl->GetSubFOlderLevel())
					{	
						if(doLog)
							m_logHelper.LogPAB(L"pRecurse = GetContactFolderCount(pSubFolder,count) START");

						pRecurse = GetContactCount(pSubFolder,count,doLog);

						if(doLog)
							m_logHelper.LogPAB(L"pRecurse = GetContactFolderCount(pSubFolder,count) END");
					}

					if(pRecurse) 
					{
						if(doLog)
							m_logHelper.LogPAB(L"RELEASE(pSubFolder) 2 START");

						RELEASE(pSubFolder);

						if(doLog)
							m_logHelper.LogPAB(L"RELEASE(pSubFolder) 2 END");

						pSubFolder = pRecurse;
						break;
					}
				}
				else
				{
					__recursionCount = 1;
				}
			} while(pSubFolder&& !m_pCtrl->IsPABContactCoutningAbort());
			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 2 START");

			RELEASE(pHierarchy);

			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pHierarchy) 2 END");

			m_pHierarchy=NULL;
		}
		// this may occur many times depending on how deep the recursion is; make sure we haven't already assigned m_pFolder
		if(pSubFolder && m_pFolder!=pSubFolder) {
			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pFolder) START");

			RELEASE(m_pFolder);

			if(doLog)
				m_logHelper.LogPAB(L"RELEASE(m_pFolder) END");

			m_pFolder=pSubFolder;
		}
	}
	catch(_com_error &e)
	{
		
		//TCHAR error[MAX_PATH];
		//e->GetErrorMessage(error,MAX_PATH);
		if(doLog)
		{
			m_logHelper.LogPAB(L"GetContactCount Exception Catch START:");
			m_logHelper.LogPAB(L"FolderName: " + strFolder );
			m_logHelper.LogPAB(e.ErrorMessage());
			m_logHelper.LogPAB(L"GetContactCount Exception Catch END:");
		}
	}
	return pSubFolder;
} 


LPMAPIFOLDER CExMapi::GetContacts(LPMAPIFOLDER pFolder,LONG& folderCount,LONG& contactCount,LONG& skipCount,LONG& contactIndex,CObArray& contactArray,BOOL doLog)
{
	__recursionCount++;
	wchar_t buffer[5];
	_itow(__recursionCount,buffer,10);
	if(doLog)
	{
		if(__recursionCount > 1)
			m_logHelper.LogPAB(L"-------------------------------------");
		m_logHelper.LogPAB(L"Recursive Call No: "+CString(buffer));
		m_logHelper.LogPAB(L"-------------------------------------");
	}
	CString strFilterBy = m_pCtrl->GetPABSerachBy();
	CString strStartWith = m_pCtrl->GetPABSerachStartWith();
	strStartWith = strStartWith.MakeLower();
	ULONG pageSize;
	m_pCtrl->get_PABPageSize(&pageSize);
	LONG pageNumber = m_pCtrl->GetPABPageNumber();
	LONG startIndex = -1;
	if(pageSize>0)
	{
		startIndex = (pageSize * pageNumber) - pageSize;
	}

	LPMAPIFOLDER pSubFolder=NULL;
	LPMAPITABLE pHierarchy;
	CString strFolder;

	try
	{
		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy); START");

		RELEASE(m_pHierarchy);
		
		if(doLog)
			m_logHelper.LogPAB(L"RELEASE(m_pHierarchy); END");

		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder); START");

		pHierarchy = GetHierarchy(pFolder);

		if(doLog)
			m_logHelper.LogPAB(L"pHierarchy = GetHierarchy(pFolder); END");

		if(pHierarchy) 
		{
			LPMAPIFOLDER pRecurse = NULL;
			do 
			{
				RELEASE(pSubFolder);
				m_pHierarchy = pHierarchy;			

				if(doLog)
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder); START");

				pSubFolder = GetNextSubFolder(strFolder,pFolder);

				if(doLog)
					m_logHelper.LogPAB(L"pSubFolder = GetNextSubFolder(strFolder,pFolder); END");
			
				if(pSubFolder) 
				{
					//if(strFolder == L"Deleted Items")
					//	continue;

					if(IsContactFolder(pSubFolder))
					{
						folderCount++;
	 					if(doLog)
							m_logHelper.LogPAB(L"if(GetContents(pSubFolder)) START");

						if(GetContents(pSubFolder))
						{
							if(doLog)
								m_logHelper.LogPAB(L"if(GetContents(pSubFolder)) END");

							SortContents(TABLE_SORT_ASCEND,PR_DISPLAY_NAME);
							//SortContents(TABLE_SORT_ASCEND,PR_SUBJECT);

							CString strText;
							CExMapiContact mapiContact;
							CString name;
							CString email;
							BOOL mustFilter = FALSE;
							mustFilter = strStartWith != "";
							
							if(doLog)
								m_logHelper.LogPAB(L"Start Reading Contact properties");

							while(GetNextContact(mapiContact) && !m_pCtrl->IsPABSearchAbort()) 
							{
								if(pageSize > 0)
								{
									if(contactCount == pageSize)
									{
										break;								
									}
								}

								contactIndex++;
								if(mustFilter)
								{
									if(strFilterBy == "Name")
									{
										mapiContact.GetName(name,PR_GIVEN_NAME);
										name = name.MakeLower();
										if(name.Find(strStartWith) != 0)				
										{
											continue;									
										}
									}

									if(strFilterBy == "Email")
									{
										mapiContact.GetEmail(email);
										email = email.MakeLower();
										if(email.Find(strStartWith) != 0)
										{
											continue;
										}
									}																
								}
								if(startIndex > 0 && skipCount < startIndex)
								{
									skipCount++;
									continue;								
								}

								CContact* pContact = new CContact();

								pContact->SetId(contactIndex);

								pContact->SetFolderName(strFolder);

								mapiContact.GetName(strText);
								pContact->SetFullName(strText);

								mapiContact.GetCompany(strText);
								pContact->SetCompany(strText);

								mapiContact.GetIMAddress(strText);
								pContact->SetIMAddress(strText);

								mapiContact.GetName(strText,PR_GIVEN_NAME);
								pContact->SetFirstName(strText);

								mapiContact.GetName(strText,PR_MIDDLE_NAME);
								pContact->SetMiddleName(strText);

								mapiContact.GetName(strText,PR_SURNAME);
								pContact->SetLastName(strText);

								mapiContact.GetEmail(strText);
								pContact->SetEmail(strText);	

								mapiContact.GetEmail(strText,2);//Email 2
								pContact->SetEmail2(strText);

								mapiContact.GetEmail(strText,3);//Email 3
								pContact->SetEmail3(strText);

								mapiContact.GetPhoneNumber(strText,PR_BUSINESS_TELEPHONE_NUMBER);
								pContact->SetBusinessPhone(strText);

								mapiContact.GetPhoneNumber(strText,PR_COMPANY_MAIN_PHONE_NUMBER);
								pContact->SetCompanyPhone(strText);

								mapiContact.GetPhoneNumber(strText,PR_BUSINESS_FAX_NUMBER);
								pContact->SetFax(strText);

								mapiContact.GetPhoneNumber(strText,PR_MOBILE_TELEPHONE_NUMBER);
								pContact->SetMobilePhone(strText);

								mapiContact.GetPhoneNumber(strText,PR_HOME_TELEPHONE_NUMBER);
								pContact->SetHomePhone(strText);

								if(m_pCtrl->FetchUnique())
								{
									BOOL isContactExist = FALSE;
									for(int i=0;i<contactArray.GetCount();i++)
									{
										CContact* pTemp = (CContact*)contactArray.GetAt(i);

										if(pContact->GetEmail() != "" && pTemp->GetEmail() == pContact->GetEmail())
										{
											isContactExist = TRUE;
											break;																	
										}
										else if(pContact->GetFullName() != "" && pTemp->GetFullName() == pContact->GetFullName())
										{
											isContactExist = TRUE;
											break;
										}
										else if(pContact->GetMobilePhone() != "" && pTemp->GetMobilePhone() == pContact->GetMobilePhone())
										{
											isContactExist = TRUE;
											break;
										}
										else if(pContact->GetHomePhone() != "" && pTemp->GetHomePhone() == pContact->GetHomePhone())
										{
											isContactExist = TRUE;
											break;
										}
										else if(pContact->GetBusinessPhone() != "" && pTemp->GetBusinessPhone() == pContact->GetBusinessPhone())
										{
											isContactExist = TRUE;
											break;
										}
									}

									if(isContactExist)
									{
										continue;
									}
								}														
								contactArray.Add(pContact);							
								contactCount++;							
							}
						}
					}

					m_pHierarchy = NULL; // so we don't release it in subsequent drilldown
					if(pageSize > 0)
					{
						if(contactCount == pageSize)
							break;								
					}
					
					if(__recursionCount < m_pCtrl->GetSubFOlderLevel())
					{
						pRecurse = GetContacts(pSubFolder,folderCount,contactCount,skipCount,contactIndex,contactArray,doLog);
					}
					if(pRecurse) 
					{
						RELEASE(pSubFolder);
						pSubFolder = pRecurse;
						break;
					}
				}
				else
				{
					__recursionCount = 1;
				}
				if(pageSize > 0)
				{
					if(contactCount == pageSize)
						break;								
				}

			} while(pSubFolder && !m_pCtrl->IsPABSearchAbort());
			RELEASE(pHierarchy);
			m_pHierarchy=NULL;
		}
		// this may occur many times depending on how deep the recursion is; make sure we haven't already assigned m_pFolder
		if(pSubFolder && m_pFolder!=pSubFolder) {
			RELEASE(m_pFolder);
			m_pFolder=pSubFolder;
		}
	}
	catch(_com_error &e)
	{
		if(doLog)
		{
			m_logHelper.LogPAB(L"GetContactFolderCount Exception Catch START:");
			m_logHelper.LogPAB(L"FolderName: " + strFolder );
			m_logHelper.LogPAB(e.ErrorMessage());
			m_logHelper.LogPAB(L"GetContactFolderCount Exception Catch END:");
		}
	}
	return pSubFolder;
} 



ULONG CExMapi::GetContactFolderCount()
{	
	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABFolderCount START:");
		m_logHelper.LogPAB(L"------------------------------------");
	}

	ULONG count = 0;
	LPMAPIFOLDER pFolder = NULL;
	LPMAPITABLE lpMsgStoreTable = NULL;
	lpMsgStoreTable = GetMessageStoresTable();	
	
	int msgStoreCount = 0;
	wchar_t buffer[5];
	while(OpenNextMessageStore(lpMsgStoreTable))
	{			
		msgStoreCount++;
		_itow(msgStoreCount,buffer,10);
		if(doLog)
		{
			if(msgStoreCount > 1)
				m_logHelper.LogPAB(L"-------------------------------------");
			m_logHelper.LogPAB(L"Message Store No : "+CString(buffer));
			m_logHelper.LogPAB(L"-------------------------------------");
		}

		//pFolder = GetContactFolderCount(OpenRootFolder(),count);			
		pFolder = OpenContacts();
		if(pFolder && IsContactFolder(pFolder))
		{			
			count++;
			pFolder = GetContactFolderCount(pFolder,count,doLog);			
		}
		if(m_pCtrl->IsPABFolderCoutningAbort())
			break;
		__recursionCount = 0;
	}
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABFolderCount END:");
		m_logHelper.LogPAB(L"------------------------------------");		
		m_logHelper.ClosePABLog();
	}
	__recursionCount = 0;
	RELEASE(lpMsgStoreTable);		
	
	return count;
}
ULONG CExMapi::GetContactCount()
{	
	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABCotactCount START:");
		m_logHelper.LogPAB(L"------------------------------------");
	}

	ULONG count = 0;
	LPMAPIFOLDER pFolder = NULL;
	LPMAPITABLE lpMsgStoreTable = NULL;
	lpMsgStoreTable = GetMessageStoresTable();	
	CString strFolder;
	int msgStoreCount = 0;
	wchar_t buffer[5];
	while(OpenNextMessageStore(lpMsgStoreTable))
	{			
		//pFolder = GetContactCount(OpenRootFolder(),count);			

		msgStoreCount++;
		_itow(msgStoreCount,buffer,10);
		if(doLog)
		{
			if(msgStoreCount > 1)
				m_logHelper.LogPAB(L"-------------------------------------");
			m_logHelper.LogPAB(L"Message Store No : "+CString(buffer));
			m_logHelper.LogPAB(L"-------------------------------------");
		}

		pFolder = OpenContacts();
		if(pFolder && IsContactFolder(pFolder))
		{
			strFolder = GetFolderName(pFolder);
			if(GetContents(pFolder))
			{
				count += GetRowCount();
			}		
			pFolder = GetContactCount(pFolder,count,doLog);
		}		
		if(m_pCtrl->IsPABContactCoutningAbort())
			break;
		__recursionCount = 0;
	}
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABCotactCount END:");
		m_logHelper.LogPAB(L"------------------------------------");	
		m_logHelper.ClosePABLog();
	}
	__recursionCount = 0;
	RELEASE(lpMsgStoreTable);		
	
	return count;
}
BOOL CExMapi::GetContacts(CObArray& contactArray,CTRError& error)
{
	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABCotacts START:");
		m_logHelper.LogPAB(L"------------------------------------");
	}

	CString strFilterBy = m_pCtrl->GetPABSerachBy();
	CString strStartWith = m_pCtrl->GetPABSerachStartWith();
	strStartWith = strStartWith.MakeLower();
	ULONG pageSize;
	m_pCtrl->get_PABPageSize(&pageSize);
	ULONG pageNumber = m_pCtrl->GetPABPageNumber();
	LONG contactCount = 0;
	LONG skipCount = 0;
	LONG startIndex = -1;
	LONG contactIndex = 0;
	LONG folderCount = 0;
	if(pageSize>0)
	{
		startIndex = (pageSize * pageNumber) - pageSize;
	}

	CString strFolder;
	int msgStoreCount = 0;
	wchar_t buffer[5];

	LPMAPIFOLDER pSubFolder=NULL;	
	LPMAPITABLE lpMsgStoreTable = NULL;
	lpMsgStoreTable = GetMessageStoresTable();	
	if(lpMsgStoreTable)
	{
		while(OpenNextMessageStore(lpMsgStoreTable))
		{	
			//GetContacts(OpenRootFolder(),folderCount,contactCount,skipCount,contactIndex,contactArray);

			msgStoreCount++;
			_itow(msgStoreCount,buffer,10);
			if(doLog)
			{
				if(msgStoreCount > 1)
					m_logHelper.LogPAB(L"-------------------------------------");
				m_logHelper.LogPAB(L"Message Store No : "+CString(buffer));
				m_logHelper.LogPAB(L"-------------------------------------");
			}

			pSubFolder = OpenContacts();
			if(pSubFolder && IsContactFolder(pSubFolder))
			{
				if(doLog)
					m_logHelper.LogPAB(L"Open Main Contacts Folder");

				strFolder = GetFolderName(pSubFolder);
				folderCount++;
				if(GetContents(pSubFolder))
				{
					SortContents(TABLE_SORT_ASCEND,PR_DISPLAY_NAME);
					//SortContents(TABLE_SORT_ASCEND,PR_SUBJECT);

					CString strText;
					CExMapiContact mapiContact;
					CString name;
					CString email;
					BOOL mustFilter = FALSE;
					mustFilter = strStartWith != "";
					
					if(doLog)
						m_logHelper.LogPAB(L"Start reading contacts from main contact fodler");

					while(GetNextContact(mapiContact) && !m_pCtrl->IsPABSearchAbort()) 
					{
						if(pageSize > 0)
						{
							if(contactCount == pageSize)
							{
								break;								
							}
						}

						contactIndex++;
						if(mustFilter)
						{
							if(strFilterBy == "Name")
							{
								mapiContact.GetName(name,PR_GIVEN_NAME);
								name = name.MakeLower();
								if(name.Find(strStartWith) != 0)				
								{
									continue;									
								}
							}

							if(strFilterBy == "Email")
							{
								mapiContact.GetEmail(email);
								email = email.MakeLower();
								if(email.Find(strStartWith) != 0)
								{
									continue;
								}
							}																
						}
						if(startIndex > 0 && skipCount < startIndex)
						{
							skipCount++;
							continue;								
						}

						CContact* pContact = new CContact();

						pContact->SetId(contactIndex);

						pContact->SetFolderName(strFolder);

						mapiContact.GetName(strText);
						pContact->SetFullName(strText);

						mapiContact.GetCompany(strText);
						pContact->SetCompany(strText);

						mapiContact.GetIMAddress(strText);
						pContact->SetIMAddress(strText);

						mapiContact.GetName(strText,PR_GIVEN_NAME);
						pContact->SetFirstName(strText);

						mapiContact.GetName(strText,PR_MIDDLE_NAME);
						pContact->SetMiddleName(strText);

						mapiContact.GetName(strText,PR_SURNAME);
						pContact->SetLastName(strText);

						mapiContact.GetEmail(strText);
						pContact->SetEmail(strText);	

						mapiContact.GetEmail(strText,2);//Email 2
						pContact->SetEmail2(strText);

						mapiContact.GetEmail(strText,3);//Email 3
						pContact->SetEmail3(strText);

						mapiContact.GetPhoneNumber(strText,PR_BUSINESS_TELEPHONE_NUMBER);
						pContact->SetBusinessPhone(strText);

						mapiContact.GetPhoneNumber(strText,PR_COMPANY_MAIN_PHONE_NUMBER);
						pContact->SetCompanyPhone(strText);

						mapiContact.GetPhoneNumber(strText,PR_BUSINESS_FAX_NUMBER);
						pContact->SetFax(strText);

						mapiContact.GetPhoneNumber(strText,PR_MOBILE_TELEPHONE_NUMBER);
						pContact->SetMobilePhone(strText);

						mapiContact.GetPhoneNumber(strText,PR_HOME_TELEPHONE_NUMBER);
						pContact->SetHomePhone(strText);

						if(m_pCtrl->FetchUnique())
						{
							BOOL isContactExist = FALSE;
							for(int i=0;i<contactArray.GetCount();i++)
							{
								CContact* pTemp = (CContact*)contactArray.GetAt(i);

								if(pContact->GetEmail() != "" && pTemp->GetEmail() == pContact->GetEmail())
								{
									isContactExist = TRUE;
									break;																	
								}
								else if(pContact->GetFullName() != "" && pTemp->GetFullName() == pContact->GetFullName())
								{
									isContactExist = TRUE;
									break;
								}
								else if(pContact->GetMobilePhone() != "" && pTemp->GetMobilePhone() == pContact->GetMobilePhone())
								{
									isContactExist = TRUE;
									break;
								}
								else if(pContact->GetHomePhone() != "" && pTemp->GetHomePhone() == pContact->GetHomePhone())
								{
									isContactExist = TRUE;
									break;
								}
								else if(pContact->GetBusinessPhone() != "" && pTemp->GetBusinessPhone() == pContact->GetBusinessPhone())
								{
									isContactExist = TRUE;
									break;
								}
							}

							if(isContactExist)
							{
								continue;
							}
						}														
						contactArray.Add(pContact);							
						contactCount++;							
					}
				}
				if(doLog)
					m_logHelper.LogPAB(L"Open Sub Contacts Folder");
				GetContacts(pSubFolder,folderCount,contactCount,skipCount,contactIndex,contactArray,doLog);
			}
			if(pageSize > 0)
			{
				if(contactCount == pageSize)
					break;								
			}
			if(m_pCtrl->IsPABSearchAbort())
				break;
			__recursionCount = 0;
		}
	}	
	if(doLog)
	{
		m_logHelper.LogPAB(L"------------------------------------");
		m_logHelper.LogPAB(L"Function GetPABCotacts END:");
		m_logHelper.LogPAB(L"------------------------------------");	
		m_logHelper.ClosePABLog();
	}
	__recursionCount = 0;
	RELEASE(lpMsgStoreTable);		

	if(folderCount == 0 && !m_pCtrl->IsPABSearchAbort())
	{
		error.code = 107;
		error.err = "TRContactFinder Cannot find private folders.";
		return FALSE;
	}

	return TRUE;
}