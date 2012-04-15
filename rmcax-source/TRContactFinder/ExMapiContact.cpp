////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiContact.cpp
// Description: MAPI Contact class wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include <MapiUtil.h>
#include "ExMapi.h"
#include "ExMapiNamedProperty.h"

/////////////////////////////////////////////////////////////
// CExMapiContact

const ULONG PhoneNumberIDs[]={
	PR_PRIMARY_TELEPHONE_NUMBER, PR_BUSINESS_TELEPHONE_NUMBER, PR_HOME_TELEPHONE_NUMBER, 
	PR_CALLBACK_TELEPHONE_NUMBER, PR_BUSINESS2_TELEPHONE_NUMBER, PR_MOBILE_TELEPHONE_NUMBER,
	PR_RADIO_TELEPHONE_NUMBER, PR_CAR_TELEPHONE_NUMBER, PR_OTHER_TELEPHONE_NUMBER,
	PR_PAGER_TELEPHONE_NUMBER, PR_PRIMARY_FAX_NUMBER, PR_BUSINESS_FAX_NUMBER,
	PR_HOME_FAX_NUMBER, PR_TELEX_NUMBER, PR_ISDN_NUMBER, PR_ASSISTANT_TELEPHONE_NUMBER,
	PR_HOME2_TELEPHONE_NUMBER, PR_TTYTDD_PHONE_NUMBER, PR_COMPANY_MAIN_PHONE_NUMBER, 0
};

const ULONG NameIDs[]={ PR_DISPLAY_NAME, PR_GIVEN_NAME, PR_MIDDLE_NAME, PR_SURNAME, PR_INITIALS, 0 };

CExMapiContact::CExMapiContact()
{
	m_pUser=NULL;
	m_pMAPI=NULL;
	m_entry.cb=0;
	SetEntryID(NULL);
}

CExMapiContact::~CExMapiContact()
{
	Close();
}

void CExMapiContact::SetEntryID(SBinary* pEntry)
{
	if(m_entry.cb) delete [] m_entry.lpb;
	m_entry.lpb=NULL;

	if(pEntry) {
		m_entry.cb=pEntry->cb;
		if(m_entry.cb) {
			m_entry.lpb=new BYTE[m_entry.cb];
			memcpy(m_entry.lpb,pEntry->lpb,m_entry.cb);
		}
	} else {
		m_entry.cb=0;
	}
}

BOOL CExMapiContact::Open(CExMapi* pMAPI,SBinary entry)
{
	Close();
	m_pMAPI=pMAPI;
	ULONG ulObjType;
	if(m_pMAPI->GetSession()->OpenEntry(entry.cb,(LPENTRYID)entry.lpb,NULL,MAPI_BEST_ACCESS,&ulObjType,(LPUNKNOWN*)&m_pUser)!=S_OK) return FALSE;

	SetEntryID(&entry);
	return TRUE;
}

void CExMapiContact::Close()
{
	SetEntryID(NULL);
	RELEASE(m_pUser);
	m_pMAPI=NULL;
}

HRESULT CExMapiContact::GetProperty(ULONG ulProperty,LPSPropValue& pProp)
{
	ULONG ulPropCount;
	ULONG p[2]={ 1,ulProperty };
	return m_pUser->GetProps((LPSPropTagArray)p, CExMapi::cm_nMAPICode, &ulPropCount, &pProp);
}

BOOL CExMapiContact::GetPropertyString(CString& strProperty,ULONG ulProperty)
{
	LPSPropValue pProp;
	if(GetProperty(ulProperty,pProp)==S_OK) {
		strProperty=CExMapi::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	} else {
		strProperty=_T("");
		return FALSE;
	}
}

BOOL CExMapiContact::GetNamedProperty(LPCTSTR szFieldName,CString& strField)
{
	CExMapiNamedProperty prop(m_pUser);
	return prop.GetNamedProperty(szFieldName,strField);
}

BOOL CExMapiContact::GetName(CString& strName,ULONG ulNameID)
{
	ULONG i=0;
	while(NameIDs[i]!=ulNameID && NameIDs[i]>0) i++;
	if(!NameIDs[i]) {
		strName=_T("");
		return FALSE;
	} else {
		return GetPropertyString(strName,ulNameID);
	}
}

int CExMapiContact::GetOutlookEmailID(int nIndex)
{
	ULONG ulProperty[]={ OUTLOOK_EMAIL1, OUTLOOK_EMAIL2, OUTLOOK_EMAIL3 };
	if(nIndex<1 || nIndex>3) return 0;
	return ulProperty[nIndex-1];
}

// uses the built in outlook email fields, OUTLOOK_EMAIL1 etc, minus 1 for ADDR_TYPE and +1 for EmailOriginalDisplayName
BOOL CExMapiContact::GetEmail(CString& strEmail,int nIndex)
{
	strEmail = L"";
	ULONG nID=GetOutlookEmailID(nIndex);
	if(!nID) return FALSE;

	LPSPropValue pProp;
	CExMapiNamedProperty prop(m_pUser);
	if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID-1,pProp)) {
		CString strAddrType=CExMapi::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID,pProp)) {
			strEmail=CExMapi::GetValidString(*pProp);
			MAPIFreeBuffer(pProp);
			if(strAddrType==_T("EX")) {
				// for EX types we use the original display name (seems to contain the appropriate data)
				if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID+1,pProp)) {
					strEmail=CExMapi::GetValidString(*pProp);
					MAPIFreeBuffer(pProp);
				}
			}
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CExMapiContact::GetPhoneNumber(CString& strPhoneNumber,ULONG ulPhoneNumberID)
{
	ULONG i=0;
	while(PhoneNumberIDs[i]!=ulPhoneNumberID && PhoneNumberIDs[i]>0) i++;
	if(!PhoneNumberIDs[i]) {
		strPhoneNumber=_T("");
		return FALSE;
	} else {
		return GetPropertyString(strPhoneNumber,ulPhoneNumberID);
	}
}

BOOL CExMapiContact::GetTitle(CString& strTitle)
{
	return GetPropertyString(strTitle,PR_TITLE);
}

BOOL CExMapiContact::GetCompany(CString& strCompany)
{
	return GetPropertyString(strCompany,PR_COMPANY_NAME);
}

BOOL CExMapiContact::GetIMAddress(CString& strIMAddress)
{
	//return GetPropertyString(strIMAddress,PR_IM_ADDRESS);
	strIMAddress = "";
	#ifndef _WIN32_WCE
	LPSPropValue pProp;
	CExMapiNamedProperty prop(m_pUser);
	if(prop.GetOutlookProperty(OUTLOOK_DATA1, OUTLOOK_IM_ADDRESS, pProp)) 
	{
		strIMAddress=CExMapi::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	}
#endif
	return FALSE;

}

