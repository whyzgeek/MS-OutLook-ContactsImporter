#ifndef __EXMAPICONTACT_H__
#define __EXMAPICONTACT_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: ExMapiContact.h
// Description: MAPI Contact class wrapper
//

class CExMapi;
class CExMapiContact;

/////////////////////////////////////////////////////////////
// CExMapiContact

class CExMapiContact : public CObject
{
public:
	CExMapiContact();
	~CExMapiContact();

	enum { OUTLOOK_DATA1=0x00062004, OUTLOOK_EMAIL1=0x8083, OUTLOOK_EMAIL2=0x8093, OUTLOOK_EMAIL3=0x80A3,OUTLOOK_IM_ADDRESS=0x8062};

// Attributes
private:
	CExMapi* m_pMAPI;
	LPMAILUSER m_pUser;
	SBinary m_entry;

// Operations
public:

	BOOL Open(CExMapi* pMAPI,SBinary entry);
	void Close();	

	BOOL GetName(CString& strName,ULONG ulNameID=PR_DISPLAY_NAME);
	BOOL GetEmail(CString& strEmail,int nIndex=1); // 1, 2 or 3 for outlook email addresses	
	BOOL GetPhoneNumber(CString& strPhoneNumber,ULONG ulPhoneNumberID);				
	BOOL GetTitle(CString& strTitle);
	BOOL GetCompany(CString& strCompany);			
	BOOL GetIMAddress(CString& strIMAddress);

private:
	void SetEntryID(SBinary* pEntry=NULL);
	SBinary* GetEntryID() { return &m_entry; }	
	int GetOutlookEmailID(int nIndex);
	BOOL GetPropertyString(CString& strProperty,ULONG ulProperty);
	BOOL GetNamedProperty(LPCTSTR szFieldName,CString& strField);	
	HRESULT GetProperty(ULONG ulProperty,LPSPropValue &pProp);
};

#endif
