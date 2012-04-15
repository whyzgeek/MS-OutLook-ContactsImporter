////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: Nk2Parser.h
// Description: Parse Outlook NK2 File Class
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

#include <afx.h>
#include <string>
#include "contact.h"
#include "CommonMethods.h"
#include "LogHelper.h"

class CContactFinderCtrl;

class CNK2Parser
{
private:
	CString _nk2FileName;
	CContactFinderCtrl* m_pCtrl;
	CLogHelper m_logHelper;

public:
	CNK2Parser(CContactFinderCtrl* pCtrl);
	~CNK2Parser(void);
	BOOL GetNK2Contacts(CString strSearchBy,CString strStartWith,BOOL bUnique,LONG lPageSize,LONG lPageNo, CObArray& contacts,CTRError& error);
	int GetNK2ContactsCount(CTRError& e);

private:			
	BOOL ExtractNK2UDataInTextFile(TCHAR* filePath,TCHAR* error);
	static BOOL IsNK2FileExist();
	static BOOL ExtractBinResource( LPCWSTR strCustomResName, int nResourceId, tstring strOutputName);
	BOOL ConvertNK2File2TextFile(TCHAR* nk2File);	
	BOOL GetNK2File(TCHAR* nk2File);
	
	tstring ProcessRow(std::vector<byte> rowData);
	tstring ProcessSMTPRow(std::vector<byte> rowData);
	tstring ProcessExchangeRow(std::vector<byte> rowData);
	tstring BytesToWideString(std::vector<byte> bytes);		
	tstring GetType(std::vector<byte> rowBytes);
	tstring ExtractEmailFromExchangeContact(std::vector<byte> rowBytes);
	BOOL IsRowEnd(BYTE* sep);
    BOOL IsWord(BYTE* sep,int& skip);
	BOOL AreThreeZeros(BYTE* sep);   		
	BOOL IsValidEmail(const tstring email);
	void Reverse(std::vector<byte>& bytes);

};
