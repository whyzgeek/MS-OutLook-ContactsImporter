#include "StdAfx.h"
using namespace std;
#include <fstream> 
#include "CommonMethods.h"
#include "registryhelper.h"
#include <msi.h>
#include <windows.h>
#include <winver.h>
#include <shlwapi.h>

CCommonMethods::CCommonMethods(void)
{
}

CCommonMethods::~CCommonMethods(void)
{
}

BOOL CCommonMethods::IsOutlookInstalled(CTRError& err)
{
	CRegistryHelper rh;
	LONG result = rh.Open(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE",KEY_READ);
	if(result == ERROR_SUCCESS)
	{
		CString strValue;
		rh.QueryValue(strValue,L"Path");
		rh.Close();
		if(strValue != "")
			return TRUE;
		else
		{
			err.code = 101;
			err.err = L"Outlook is not installed.";
		}
	}
	else if(result == ERROR_ACCESS_DENIED)
	{
		err.code = 103;
		err.err = L"Access is denied. Cannot initialize TRContactFinder.";
	}
	else if(result == ERROR_FILE_NOT_FOUND)
	{
		err.code = 101;
		err.err = L"Outlook is not installed.";
	}
	return FALSE;
}
BOOL CCommonMethods::IsVaildOutlookVersion(CTRError& err)
{
	CRegistryHelper rh;
	LONG result = rh.Open(HKEY_LOCAL_MACHINE,_TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE"),KEY_READ);
	if(result == ERROR_SUCCESS)
	{		
		CString strValue;
		rh.QueryValue(strValue,L"Path");
		rh.Close();
		if(strValue != "")
		{					
			strValue += L"Outlook.exe";
			WCHAR wFile[_MAX_PATH];
			memset(wFile,'\0',_MAX_PATH);
			wcscpy(wFile,strValue.AllocSysString());
			DWORD dwDummy;
			DWORD dwFVISize = GetFileVersionInfoSize(wFile , &dwDummy );
			LPBYTE lpVersionInfo = new BYTE[dwFVISize]; 
			if(GetFileVersionInfo(wFile,0,dwFVISize,lpVersionInfo))
			{
				VS_FIXEDFILEINFO *lpFfi;
				UINT uLen; 
				VerQueryValue( lpVersionInfo , _T("\\") , (LPVOID *)&lpFfi , &uLen );
				DWORD dwFileVersionMS = lpFfi->dwFileVersionMS;
				DWORD dwFileVersionLS = lpFfi->dwFileVersionLS;
				delete [] lpVersionInfo;
				DWORD dwLeftMost = HIWORD(dwFileVersionMS);
				DWORD dwSecondLeft = LOWORD(dwFileVersionMS);
				DWORD dwSecondRight = HIWORD(dwFileVersionLS);
				DWORD dwRightMost = LOWORD(dwFileVersionLS);
				if(dwLeftMost == 11 || dwLeftMost == 12)
				{
					return TRUE;
				}
			}

		}
		else
		{
			err.code = 106;
			err.err = L"TRContactFinder only support Outlook 2003 and Outlook 2007.";
		}
	}
	else if(result == ERROR_ACCESS_DENIED)
	{
		err.code = 103;
		err.err = L"Access is denied. Cannot initialize TRContactFinder.";
	}
	else if(result == ERROR_FILE_NOT_FOUND)
	{
		err.code = 101;
		err.err = L"Outlook is not installed.";
	}
	return FALSE;
}

BOOL CCommonMethods::IsOutlookDefaultMailClient(CTRError& err)
{
	CRegistryHelper rh;
	LONG result = rh.Open(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Clients\\Mail",KEY_READ);
	if(result == ERROR_SUCCESS)
	{
		CString strValue;
		rh.QueryValue(strValue,NULL,FALSE);
		rh.Close();
		if(strValue != "" && strValue.MakeLower() == "microsoft outlook")
			return TRUE;
		else
		{
			err.code = 102;
			//err.err = L"Outlook is not set as default mail client.\r\nTo set do following settings in IE:\r\n IE -> Tools -> Internet Options -> Programs Tab -> E-mail -> Microsoft Outlook.";
			err.err = L"Outlook is not set as default mail client.";
		}
	}
	else if(result == ERROR_ACCESS_DENIED)
	{
		err.code = 103;
		err.err = L"Access is denied. Cannot initialize TRContactFinder.";
	}
	return FALSE;
}

HRESULT CCommonMethods::GetBrowserVersion(LPDWORD pdwMajor, LPDWORD pdwMinor, LPDWORD pdwBuild)
{
    HINSTANCE   hBrowser;

    if(IsBadWritePtr(pdwMajor, sizeof(DWORD))
        || IsBadWritePtr(pdwMinor, sizeof(DWORD))
        || IsBadWritePtr(pdwBuild, sizeof(DWORD)))
        return E_INVALIDARG;

    *pdwMajor = 0;
    *pdwMinor = 0;
    *pdwBuild = 0;

    //Load the DLL.

    hBrowser = LoadLibrary(TEXT("shdocvw.dll"));

    if(hBrowser) 
    {

        HRESULT  hr = S_OK;
        DLLGETVERSIONPROC pDllGetVersion;
		char  proc[14] = "DllGetVersion";
		proc[13] = '\0';
        pDllGetVersion = (DLLGETVERSIONPROC)GetProcAddress(hBrowser, proc);

        if(pDllGetVersion) 
        {

            DLLVERSIONINFO    dvi;
            ZeroMemory(&dvi, sizeof(dvi));
            dvi.cbSize = sizeof(dvi);
            hr = (*pDllGetVersion)(&dvi);

            if(SUCCEEDED(hr)) 
            {
                *pdwMajor = dvi.dwMajorVersion;
                *pdwMinor = dvi.dwMinorVersion;
                *pdwBuild = dvi.dwBuildNumber;
            }

        } 
        else 
        {
            //If GetProcAddress failed, there is a problem 

            // with the DLL.


            hr = E_FAIL;
        }
        FreeLibrary(hBrowser);
        return hr;
    }
    return E_FAIL;
}



BOOL CCommonMethods::IsValidIE()
{
	DWORD major;
	DWORD minor;
	DWORD build;
	if(SUCCEEDED(CCommonMethods::GetBrowserVersion(&major,&minor,&build)))
	{
		if(major < 6)// Less than IE6
			return FALSE;
	}
	return TRUE;
}



BOOL CCommonMethods::IsFileExist(LPCTSTR filePath)
{
	CFileFind fileFind;
	return fileFind.FindFile(filePath);
}
BOOL CCommonMethods::IsDirExist(LPCTSTR path, LPCTSTR dirName,BOOL bCreate)
{
	tstring searchPath = tstring(path) + L"*.*";
	
	CFileFind fileFind;
	BOOL bWorking = fileFind.FindFile(searchPath.c_str());
	BOOL bFound = FALSE;
	while(bWorking)
	{
		bWorking = fileFind.FindNextFileW();
		if(fileFind.IsDirectory())
		{
			if(fileFind.GetFileName() == CString(dirName))
			{
				bFound = TRUE;			
				break;
			}
		}
	}
	if(!bFound && bCreate)
	{		
		tstring folder = tstring(path)+ tstring(dirName) + L"\\";
		bFound = CreateDirectory(folder.c_str(),NULL);
	}
	return bFound;
}
tstring CCommonMethods::GetTempDirPath()
{
	TCHAR tempPath[_MAX_PATH];
	memset(tempPath,'\0',_MAX_PATH);
	GetTempPath(_MAX_PATH,tempPath);	
	
	return tstring(tempPath);
}

tstring CCommonMethods::GetLocalAppDataFolder()
{	
	WCHAR appDataFolder[MAX_PATH];
	SHGetSpecialFolderPath(NULL,appDataFolder,CSIDL_LOCAL_APPDATA,FALSE);	
	return tstring(appDataFolder);
}


BOOL CCommonMethods::GetFileData(tstring strNK2FileName,std::vector<byte>& fileData)
{
    if(strNK2FileName.empty())
        return FALSE;

    fileData.clear();	
	std::ifstream file;
	try
	{				
		file.open(strNK2FileName.c_str(), ios_base::binary);
		if(!file.is_open())
		{
			return FALSE;
		}
		
	}
	catch(exception ex)
	{
		return FALSE;
	}

	// get the length of the file
	file.seekg(0, ios::end);
	size_t fileSize = file.tellg();
	file.seekg(0, ios::beg);

	// create a vector to hold all the bytes in the file	
	fileData.resize(fileSize);

	// read the file
	file.read(reinterpret_cast<char*>(&fileData[0]), fileSize);

	// close the file
	file.close();	

    return TRUE;
}







void CCommonMethods::Tokenize(const std::wstring& str, std::vector<std::wstring>& tokens, const std::wstring& delimiters)
{
	// Skip delimiters at beginning.
	std::wstring::size_type lastPos = str.find_first_not_of(delimiters, 0);
	// Find first "non-delimiter".
	std::wstring::size_type pos     = str.find_first_of(delimiters, lastPos);
	
	while (std::wstring::npos != pos || std::wstring::npos != lastPos)
	{
		// Found a token, add it to the vector.
		tokens.push_back(str.substr(lastPos, pos - lastPos));
		// Skip delimiters.  Note the "not_of"
		lastPos = str.find_first_not_of(delimiters, pos);
		// Find next "non-delimiter"
		pos = str.find_first_of(delimiters, lastPos);
	}
}