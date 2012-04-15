////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: Contact.cpp
// Description: Contact Properties Wrapper
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Contact.h"

// CContact

CContact::CContact()
{
	m_nId = -1;
	m_strEmail = "";
	m_strComapny = "";
	m_strFullName = "";
	m_strFirstName = "";
	m_strMiddleName = "";
	m_strLastName = "";	
	m_strHomePhone = "";
	m_strCompanyPhone = "";
	m_strBusinessPhone = "";
	m_strMobilePhone = "";
	m_strFax = "";
	m_strIMAddress = "";
}

CContact::~CContact()
{
}

void CContact::SetId(int nId)
{	
	m_nId = nId;
}
void CContact::SetFolderName(CString strFolderName)
{	
	m_strFolderName = strFolderName;
}
void CContact::SetEmail(CString strEmail)
{	
	m_strEmail = strEmail;
}
void CContact::SetEmail2(CString strEmail2)
{	
	m_strEmail2 = strEmail2;
}
void CContact::SetEmail3(CString strEmail3)
{	
	m_strEmail3 = strEmail3;
}
void CContact::SetCompany(CString strCompany)
{	
	m_strComapny = strCompany;
}
void CContact::SetFullName(CString strFullName)
{	
	m_strFullName = strFullName;
}
void CContact::SetIMAddress(CString strIMAddress)
{	
	m_strIMAddress = strIMAddress;
}
void CContact::SetFirstName(CString strFirstName)
{
	m_strFirstName = strFirstName;
}
void CContact::SetMiddleName(CString strMiddleName)
{	
	m_strMiddleName = strMiddleName;
}
void CContact::SetLastName(CString strLastName)
{	
	m_strLastName = strLastName;
}
void CContact::SetBusinessPhone(CString strBusinessPhone)
{	
	m_strBusinessPhone = strBusinessPhone;
}
void CContact::SetHomePhone(CString strHomePhone)
{	
	m_strHomePhone = strHomePhone;
}
void CContact::SetCompanyPhone(CString strCompanyPhone)
{	
	m_strCompanyPhone = strCompanyPhone;
}
void CContact::SetMobilePhone(CString strMobilePhone)
{	
	m_strMobilePhone = strMobilePhone;
}
void CContact::SetFax(CString strFax)
{	
	m_strFax = strFax;
}

CString CContact::GetFolderName()
{
	return m_strFolderName;
}
CString CContact::GetEmail()
{
	return m_strEmail;
}
CString CContact::GetEmail2()
{
	return m_strEmail2;
}
CString CContact::GetEmail3()
{
	return m_strEmail3;
}
CString CContact::GetCompany()
{
	return m_strComapny;
}
CString CContact::GetFullName()
{
	return m_strFullName;
}
CString CContact::GetFirstName()
{
	return m_strFirstName;
}
CString CContact::GetMiddleName()
{
	return m_strMiddleName;
}
CString CContact::GetLastName()
{
	return m_strLastName;
}
CString CContact::GetIMAddress()
{
	return m_strIMAddress;
}
CString CContact::GetBusinessPhone()
{
	return m_strBusinessPhone;
}
CString CContact::GetHomePhone()
{
	return m_strHomePhone;
}
CString CContact::GetCompanyPhone()
{
	return m_strCompanyPhone;
}
CString CContact::GetMobilePhone()
{
	return m_strMobilePhone;
}
CString CContact::GetFax()
{
	return m_strFax;
}
CString CContact::ToJsonString()
{
	CString json;

	json += L"{";

	char buffer [5];
	_itoa(m_nId,buffer,10);
	
	json += L"\"id\":" + CString(buffer) + L",";	
	json += L"\"folder\":\"" + ReplaceSpecialChracters(m_strFolderName) + L"\",";	
	json += L"\"fullName\":\"" + ReplaceSpecialChracters(m_strFullName) + L"\",";
	json += L"\"firstName\":\"" + ReplaceSpecialChracters(m_strFirstName) + L"\",";
	json += L"\"middleName\":\"" + ReplaceSpecialChracters(m_strMiddleName) + L"\",";
	json += L"\"lastName\":\"" + ReplaceSpecialChracters(m_strLastName) + L"\",";
	json += L"\"company\":\"" + ReplaceSpecialChracters(m_strComapny) + L"\",";	
	json += L"\"imAddress\":\"" + ReplaceSpecialChracters(m_strIMAddress) + L"\",";	

	json += L"\"emails\":[" ;
	json += L"\"" + ReplaceSpecialChracters(m_strEmail) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strEmail2) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strEmail3) + L"\"],";

	json += L"\"phones\":[" ;
	json += L"\"" + ReplaceSpecialChracters(m_strBusinessPhone) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strCompanyPhone) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strHomePhone) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strFax) + L"\",";
	json += L"\"" + ReplaceSpecialChracters(m_strMobilePhone) + L"\"]";
	
	json += L"}";

	return json;
}

CString CContact::ReplaceSpecialChracters(CString str)
{
	CString result;
	for(int i=0; i<str.GetLength();i++)
	{
		CString temp = CString(str.GetAt(i));
		if(temp == L"\"")
			temp.Replace(CString("\""),CString("\\\""));//double quotes
		if(temp == "\\")
			temp.Replace(CString("\\"),CString("\\\\"));//balck slash

		result += temp;
	}
	return result;
}