////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiNamedProperty.cpp
// Description: MAPI Named Property class wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include <MapiUtil.h>
#include "ExMapi.h"
#include "ExMapiNamedProperty.h"

/////////////////////////////////////////////////////////////
// CExMapiNamedProperty

CExMapiNamedProperty::CExMapiNamedProperty(IMAPIProp* pItem)
{
	m_pItem=pItem;
}

CExMapiNamedProperty::~CExMapiNamedProperty()
{
}

// gets a custom outlook property (ie EmailAddress1 of a contact)
BOOL CExMapiNamedProperty::GetOutlookProperty(ULONG ulData1,ULONG ulProperty,LPSPropValue& pProp)
{
	if(!m_pItem) return FALSE;
	const GUID guidOutlookEmail1={ulData1, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&guidOutlookEmail1;
	nameID.ulKind=MNID_ID;
	nameID.Kind.lID=ulProperty;

	LPMAPINAMEID lpNameID[1]={ &nameID };
	LPSPropTagArray lppPropTags; 

	HRESULT hr=E_INVALIDARG;
	if(m_pItem->GetIDsFromNames(1,lpNameID,0,&lppPropTags)==S_OK) {
		ULONG ulPropCount;
		hr=m_pItem->GetProps(lppPropTags, CExMapi::cm_nMAPICode, &ulPropCount, &pProp);
		MAPIFreeBuffer(lppPropTags);
	}
	return (hr==S_OK);
}

const GUID GUIDPublicStrings={0x00020329, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

BOOL CExMapiNamedProperty::GetNamedProperty(LPCTSTR szFieldName,LPSPropValue &pProp)
{
	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&GUIDPublicStrings;
	nameID.ulKind=MNID_STRING;
#ifdef UNICODE
	nameID.Kind.lpwstrName=(LPWSTR)szFieldName;
#else
	WCHAR wszFieldName[256];
	MultiByteToWideChar(CP_ACP,0,szFieldName,-1,wszFieldName,255);
	nameID.Kind.lpwstrName=wszFieldName;
#endif

	LPMAPINAMEID lpNameID[1]={ &nameID };
	LPSPropTagArray lppPropTags;

	HRESULT hr=E_INVALIDARG;
	if(m_pItem->GetIDsFromNames(1,lpNameID,0,&lppPropTags)==S_OK) {
		ULONG ulPropCount;
		hr=m_pItem->GetProps(lppPropTags, CExMapi::cm_nMAPICode, &ulPropCount, &pProp);
		MAPIFreeBuffer(lppPropTags);
	}
	return (hr==S_OK);
}

BOOL CExMapiNamedProperty::GetNamedProperty(LPCTSTR szFieldName,CString& strField)
{
	LPSPropValue pProp;
	if(GetNamedProperty(szFieldName,pProp)) {
		strField=CExMapi::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	}
	return FALSE;
}

