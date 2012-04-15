#ifndef __EXMAPINAMEDPROPERTY_H__
#define __EXMAPINAMEDPROPERTY_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiNamedProperty.h
// Description: MAPI Named Property class wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////
// CExMapiNamedProperty

class CExMapiNamedProperty : public CObject
{
public:
	CExMapiNamedProperty(IMAPIProp* pItem);
	~CExMapiNamedProperty();

// Attributes
protected:
	IMAPIProp* m_pItem;

// Operations
public:
	BOOL GetOutlookProperty(ULONG ulData1,ULONG ulProperty,LPSPropValue& pProp);
	BOOL GetNamedProperty(LPCTSTR szFieldName,LPSPropValue &pProp);
	BOOL GetNamedProperty(LPCTSTR szFieldName,CString& strField);	
};

#endif
