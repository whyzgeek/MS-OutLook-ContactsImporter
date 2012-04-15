////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: Contact.h
// Description: Contact Property Class
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

// CContact command target

class CContact : public CObject
{

private:	
	int m_nId;
	CString m_strFolderName;
	CString m_strFullName;
	CString m_strEmail;
	CString m_strEmail2;
	CString m_strEmail3;
	CString m_strIMAddress;
	CString m_strComapny;
	CString m_strFirstName;
	CString m_strMiddleName;
	CString m_strLastName;
	CString m_strBusinessPhone;
	CString m_strHomePhone;
	CString m_strCompanyPhone;
	CString m_strMobilePhone;
	CString m_strFax;

public:
	CContact();
	virtual ~CContact();
	void SetId(int nId);
	void SetFolderName(CString strFolderName);
	void SetEmail(CString strEmail);
	void SetEmail2(CString strEmail2);
	void SetEmail3(CString strEmail3);
	void SetCompany(CString strCompany);
	void SetFullName(CString strFullName);
	void SetIMAddress(CString strJobTitle);
	void SetFirstName(CString strFirstName);
	void SetMiddleName(CString strMiddleName);
	void SetLastName(CString strLastName);
	void SetBusinessPhone(CString strBusinessPhone);
	void SetHomePhone(CString strHomePhone);
	void SetCompanyPhone(CString strCompanyPhone);
	void SetMobilePhone(CString strMobilePhone);
	void SetFax(CString strFax);
	
	CString GetFolderName();
	CString GetEmail();
	CString GetEmail2();
	CString GetEmail3();
	CString GetCompany();
	CString GetFullName();
	CString GetIMAddress();
	CString GetFirstName();
	CString GetMiddleName();
	CString GetLastName();
	CString GetBusinessPhone();
	CString GetHomePhone();
	CString GetCompanyPhone();
	CString GetMobilePhone();
	CString GetFax();

	CString ToJsonString();
	
private:		
	CString ReplaceSpecialChracters(CString str);
};


