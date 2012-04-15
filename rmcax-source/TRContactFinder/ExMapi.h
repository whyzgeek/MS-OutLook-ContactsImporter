#ifndef __EXMAPI_H__
#define __EXMAPI_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapi.h
// Description: Extended MAPI class 
// Works in Windows XP and Windows Vista with Outlook 2000 and 2003
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#ifdef _WIN32_WCE
#include <cemapi.h>
#include <objidl.h>
#else
#include <mapix.h>
#include <objbase.h>
#endif

#define RELEASE(s) if(s!=NULL) { s->Release();s=NULL; }

#define MAPI_NO_CACHE ((ULONG)0x00000200)
#define PR_IPM_CONTACT_ENTRYID (PROP_TAG(PT_BINARY,0x36D1))

#include "ExMapiContact.h"
#include "LogHelper.h"

////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CExMapi

class CContactFinderCtrl;

class CExMapi 
{
public:
	CExMapi(CContactFinderCtrl* pCtrl);
	~CExMapi();

	enum { PROP_MESSAGE_FLAGS, PROP_ENTRYID, MESSAGE_COLS };

// Attributes
public:
	static int cm_nMAPICode;

private:
	IMAPISession* m_pSession;
	LPMDB m_pMsgStore;
	LPMAPIFOLDER m_pFolder;
	LPMAPITABLE m_pHierarchy;
	LPMAPITABLE m_pContents;
	LPMAPITABLE m_pMsgStoresTable;
	ULONG m_sink;
	ULONG m_ulMDBFlags;
	CContactFinderCtrl* m_pCtrl;
	CLogHelper m_logHelper;
	
// Operations
public:
	static BOOL Init(BOOL bMultiThreadedNotifcations=TRUE);
	static void Term();

	BOOL Login(LPCTSTR szProfileName=NULL);
	void Logout();
	BOOL IsLogin();
	IMAPISession* GetSession() { return m_pSession; }
	LPMDB GetMessageStore() { return m_pMsgStore; }
	LPMAPIFOLDER GetFolder() { return m_pFolder; }
	void Release();
		
	ULONG GetContactFolderCount();
	ULONG GetContactCount();
	BOOL GetContacts(CObArray& contactArray,CTRError& error);

	static LPCTSTR GetValidString(SPropValue& prop);

private:
	LPCTSTR GetProfileName();

	LPMAPITABLE GetMessageStoresTable();
	LPMAPITABLE GetHierarchy(LPMAPIFOLDER pFolder=NULL);

	ULONG GetMessageStoreSupport();

	// use bInternal to specify that MAPIEx keeps track of and subsequently RELEASEs the folder 
	// remember to eventually RELEASE returned folders if calling with bInternal=FALSE
	LPMAPIFOLDER OpenFolder(unsigned long ulFolderID,BOOL bInternal);
	LPMAPIFOLDER OpenSpecialFolder(unsigned long ulFolderID,BOOL bInternal);
	LPMAPIFOLDER OpenRootFolder(BOOL bInternal=TRUE);
	LPMAPIFOLDER OpenInbox(BOOL bInternal=TRUE);			
	LPMAPIFOLDER OpenContacts(BOOL bInternal=TRUE);
	
	LPMAPIFOLDER GetNextSubFolder(CString& strFolder,LPMAPIFOLDER pFolder=NULL);
	LPMAPIFOLDER OpenSubFolder(LPCTSTR szSubFolder,LPMAPIFOLDER pFolder=NULL);
	
	LPMAPIFOLDER GetContactFolderCount(LPMAPIFOLDER pFolder,ULONG& count,BOOL doLog);
	LPMAPIFOLDER GetContactCount(LPMAPIFOLDER pFolder,ULONG& count,BOOL doLog);
	LPMAPIFOLDER GetContacts(LPMAPIFOLDER pFolder,LONG& folderCount,LONG& contactCount,LONG& skipCount,LONG& contactIndex,CObArray& contactArray,BOOL doLog);	

	BOOL OpenDefaultMessageStore(LPCTSTR szStore=NULL,ULONG ulFlags=MAPI_BEST_ACCESS/*MAPI_MODIFY | MAPI_NO_CACHE*/);
	BOOL OpenNextMessageStore(LPMAPITABLE lpMsgStoreTable,ULONG ulFlags=MAPI_BEST_ACCESS/*MAPI_MODIFY | MAPI_NO_CACHE*/);
	BOOL GetContents(LPMAPIFOLDER pFolder=NULL);
	BOOL SortContents(ULONG ulSortParam=TABLE_SORT_ASCEND,ULONG ulSortField=PR_MESSAGE_DELIVERY_TIME);
	BOOL RestrictContents(LPCTSTR pszSearchBy,LPCTSTR pszStartWith);
	BOOL GetNextContact(CExMapiContact& contact);
	BOOL AreSame(LPMAPIPROP lpMapiProp1,LPMAPIPROP lpMapiProp2);
	BOOL IsContactFolder(LPMAPIFOLDER lpFolder);
	BOOL Notify(LPNOTIFCALLBACK lpfnCallback,LPVOID lpvContext,ULONG ulEventMask=fnevNewMail);		
	
	int GetRowCount();
	CString GetFolderName(LPMAPIFOLDER lpFolder);

	static void GetNarrowString(SPropValue& prop,CString& strNarrow);

};

#endif
