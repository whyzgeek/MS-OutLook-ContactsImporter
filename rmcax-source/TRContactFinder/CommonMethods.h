////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: CommonMethods.h
// Description: Common Static Methods
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once
#include <string>
#include <vector>

typedef std::basic_string<TCHAR, std::char_traits<TCHAR>> tstring;



class CCommonMethods
{
public:
	CCommonMethods(void);
	~CCommonMethods(void);
	
	static BOOL IsOutlookInstalled(CTRError& err);
	static BOOL IsVaildOutlookVersion(CTRError& err);
	static BOOL IsOutlookDefaultMailClient(CTRError& err);
	static BOOL IsValidIE();
	static BOOL IsFileExist(LPCTSTR filePath);
	static BOOL IsDirExist(LPCTSTR path,LPCTSTR dirName,BOOL bCreate=TRUE);
	static tstring GetTempDirPath();
	static tstring GetLocalAppDataFolder();	
	static BOOL GetFileData(tstring strNK2FileName,std::vector<byte>& fileData);
	static void Tokenize(const std::wstring& str, std::vector<std::wstring>& tokens, const std::wstring& delimiters = L"|");
	
private:
	static HRESULT GetBrowserVersion(LPDWORD pdwMajor, LPDWORD pdwMinor, LPDWORD pdwBuild);

};
