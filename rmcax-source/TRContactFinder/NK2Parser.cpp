////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: Nk2Parser.cpp
// Description: Parse Outlook NK2 File 
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include <windows.h>
using namespace std;
#include <fstream> 
#include <iostream>
#include "nk2parser.h"
#include "resource.h"
#include "ContactFinderCtrl.h"
#include "textfileio.h"
#include <locale.h>



CNK2Parser::CNK2Parser(CContactFinderCtrl* pCtrl)
{	
	m_pCtrl = pCtrl;
	_nk2FileName = L"";
	if(!IsNK2FileExist())
	{
		WCHAR appDataFolder[MAX_PATH];
		if(SHGetSpecialFolderPath(NULL,appDataFolder,CSIDL_APPDATA,FALSE))
		{
			WCHAR nk2filePath[MAX_PATH];
			wcscpy(nk2filePath,appDataFolder);
			wcscat(nk2filePath,L"\\Microsoft\\Outlook\\*.nk2");

			WCHAR nk2filePath2[MAX_PATH];
			wcscpy(nk2filePath2,appDataFolder);
			wcscat(nk2filePath2,L"\\Roaming\\Microsoft\\Outlook\\*.nk2");

			CFileFind ff;
			if(ff.FindFile(nk2filePath))
			{
				ff.FindNextFile();
				CString nk2FileName = CString(appDataFolder) + L"\\Microsoft\\Outlook\\" + ff.GetFileName();

				_nk2FileName = CString(appDataFolder) + L"\\Microsoft\\Outlook\\outlook.nk2";
				
				CopyFile(nk2FileName,_nk2FileName,TRUE);

			}
			else if(ff.FindFile(nk2filePath2))
			{
				ff.FindNextFile();
				CString nk2FileName = CString(appDataFolder) + L"\\Roaming\\Microsoft\\Outlook\\" + ff.GetFileName();

				_nk2FileName = CString(appDataFolder) + L"\\Roaming\\Microsoft\\Outlook\\outlook.nk2";
				
				CopyFile(nk2FileName,_nk2FileName,TRUE);
			}
		}
	}
}

CNK2Parser::~CNK2Parser(void)
{
	if(_nk2FileName != "" && CCommonMethods::IsFileExist(_nk2FileName))
	{
		CFile::Remove(_nk2FileName);
		_nk2FileName = "";
	}
}

BOOL CNK2Parser::GetNK2Contacts(CString strSearchBy,CString strStartWith,BOOL bUnique,LONG lPageSize,LONG lPageNo, CObArray& contacts,CTRError& error)
{		
	LONG contactCount = 0;
	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogNK2(L"---------------------------------");
		m_logHelper.LogNK2(L"Function GetNK2Contacts START");
		m_logHelper.LogNK2(L"---------------------------------");
	}
	TCHAR nk2TextFile[_MAX_PATH];
	memset(nk2TextFile,'\0',_MAX_PATH);
	TCHAR terror[128];
	memset(terror,'\0',128);
	int contactIndex = 0;
	try
	{
		if(doLog)
		{		
			m_logHelper.LogNK2(L"Convert NK2 file to text file");		
		}
		//BOOL result = ExtractNK2UDataInTextFile(nk2TextFile,terror);
		BOOL result = ConvertNK2File2TextFile(nk2TextFile);
		if(!result)
		{
			//AfxMessageBox(error);
			if(doLog)
			{		
				m_logHelper.LogNK2(L"NK2 file does not exist.");		
				m_logHelper.LogNK2(L"---------------------------------");
				m_logHelper.LogNK2(L"Function GetNK2Contacts END");
				m_logHelper.LogNK2(L"---------------------------------");
				m_logHelper.CloseNK2Log();
			}
			error.code = 108;
			error.err = CString("TRContactFinder cannot find NK2 file.");
			return FALSE;
		}

		if(doLog)
		{		
			m_logHelper.LogNK2(L"Start reading NK2 text file.");		
		}

		CTextFileIO tf(nk2TextFile);
		wstring line;
		if(tf.IsValidate())
		{
			LONG skipCount = 0;
			LONG startIndex = -1;
			if(lPageSize>0)
			{
				startIndex = (lPageSize * lPageNo) - lPageSize;
			}
			while(tf.ReadLineW(line))
			{						
				size_t i=0;
				// parse the row		
				BOOL validRecord = FALSE;
				CString displayName;
				CString email;
				CString type;

				std::vector<std::wstring> tokens;
				CCommonMethods::Tokenize(line,tokens,L"\t");
				for(size_t i=0;i<tokens.size();i++)
				{				
					wstring field = tokens.at(i);

					if(i==0)	
					{
						displayName = field.c_str();										
					}
					if(i==1)				
					{
						email = field.c_str();	
						/*if(email == "MAPIDL")
						{
							type = "MAPIDL";
							email = "";
						}*/						
					}
					if(i==2 /*&& type != ""*/)				
						type = field.c_str();				

					if(i==2)
						break;										
				}

				if(type.MakeLower().Trim() == "smtp" || type.MakeLower().Trim() == "exchange")
				{			
					contactIndex++;
					if(startIndex > 0 && skipCount < startIndex)
					{
						skipCount++;
						continue;								
					}

					if(strStartWith != "")
					{
						if(strSearchBy == "Name")
						{
							if( displayName.MakeLower().Find(strStartWith.MakeLower(),0) == 0)
								validRecord = TRUE;
						}
						else if(strSearchBy == "Email")
						{
							if( email.MakeLower().Find(strStartWith.MakeLower(),0) == 0)
								validRecord = TRUE;
						}
					}
					else
					{
						validRecord = TRUE;
					}
				}

				if(!validRecord)
					continue;

				if(bUnique)
				{
					BOOL isContactExist = FALSE;
					for(int i=0;i<contacts.GetCount();i++)
					{
						CContact* pTemp = (CContact*)contacts.GetAt(i);

						if(email != "" && pTemp->GetEmail() == email)
						{
							isContactExist = TRUE;
							break;																	
						}
						else if(displayName != "" && pTemp->GetFullName() == displayName)
						{
							isContactExist = TRUE;
							break;
						}					
					}

					if(isContactExist)
						continue;
				}				
				CContact* pContact = new CContact();
				pContact->SetId(contactIndex);
				pContact->SetFullName(displayName);
				pContact->SetEmail(email);
				contacts.Add(pContact);		
				contactCount++;

				if(lPageSize > 0)
				{
					if(contactCount == lPageSize)
					break;

				}		
			}
			if(tf.Close())
			{
				CFile::Remove(nk2TextFile);
			}
		}
		if(doLog)
		{
			wchar_t buffer[5];
			_itow(contactCount,buffer,10);
			m_logHelper.LogNK2(L"NK2 Contacts Read = " + CString(buffer));
			m_logHelper.LogNK2(L"END reading NK2 text file.");		
			m_logHelper.LogNK2(L"---------------------------------");
			m_logHelper.LogNK2(L"Function GetNK2Contacts END");
			m_logHelper.LogNK2(L"---------------------------------");
			m_logHelper.CloseNK2Log();
		}
	}
	catch(CException &e)
	{
		if(doLog)
		{
			m_logHelper.LogNK2(L"GetNK2Contacts Exception Catch START:");		
			TCHAR err[MAX_PATH];
			e.GetErrorMessage(err,MAX_PATH);
			m_logHelper.LogNK2(err);
			m_logHelper.LogNK2(L"GetNK2Contacts Exception Catch END:");
			m_logHelper.CloseNK2Log();
		}
	}
	return TRUE;
}

int CNK2Parser::GetNK2ContactsCount(CTRError& e)
{	
	int count = 0;
	BOOL doLog = m_pCtrl->IsEnableLog();
	if(doLog)
	{
		m_logHelper.LogNK2(L"---------------------------------");
		m_logHelper.LogNK2(L"Function GetNK2ContactsCount START");
		m_logHelper.LogNK2(L"---------------------------------");
	}
	try
	{
		TCHAR nk2TextFile[_MAX_PATH];
		memset(nk2TextFile,'\0',_MAX_PATH);
		TCHAR error[128];
		memset(error,'\0',128);

		if(doLog)
		{		
			m_logHelper.LogNK2(L"Convert NK2 file to text file");		
		}
		//BOOL result = ExtractNK2UDataInTextFile(nk2TextFile,error);
		BOOL result = ConvertNK2File2TextFile(nk2TextFile);
		if(!result)
		{
			if(doLog)
			{		
				m_logHelper.LogNK2(L"NK2 file does not exist.");		
				m_logHelper.LogNK2(L"---------------------------------");
				m_logHelper.LogNK2(L"Function GetNK2ContactsCount END");
				m_logHelper.LogNK2(L"---------------------------------");
				m_logHelper.CloseNK2Log();
			}
			e.code = 108;
			e.err = CString("TRContactFinder cannot find NK2 file.");
			return 0;
		}
		if(doLog)
		{		
			m_logHelper.LogNK2(L"Start reading NK2 text file.");		
		}
		CTextFileIO tf(nk2TextFile);
		wstring line;
		if(tf.IsValidate())
		{
			while(tf.ReadLineW(line))
			{
				int i=0;
				// parse the row		
				BOOL validRecord = FALSE;

				wstring::size_type lastPos = line.find_first_not_of(_T("\t"), 0);
				wstring::size_type pos = line.find_first_of(_T("\t"),lastPos);

				while (wstring::npos != pos || wstring::npos != lastPos)
				{
					wstring field = line.substr(lastPos, pos - lastPos);		

					i++;

					if(i == 3)
					{
						if(CString(field.c_str()).MakeLower().Trim() == "smtp" ||
							CString(field.c_str()).MakeLower().Trim() == "exchange")
						{		
							validRecord = TRUE;
						}
						break;
					}
					
					lastPos = line.find_first_not_of(_T("\t"), pos);
					pos = line.find_first_of(_T("\t"), lastPos);
				}

				if(validRecord)
					count++;			
			}
			if(tf.Close())
			{
				CFile::Remove(nk2TextFile);
			}
		}
		if(doLog)
		{
			wchar_t buffer[5];
			_itow(count,buffer,10);
			m_logHelper.LogNK2(L"NK2 Contacts Read = " + CString(buffer));
			m_logHelper.LogNK2(L"END reading NK2 text file.");		
			m_logHelper.LogNK2(L"---------------------------------");
			m_logHelper.LogNK2(L"Function GetNK2ContactsCount END");
			m_logHelper.LogNK2(L"---------------------------------");
			m_logHelper.CloseNK2Log();
		}
	}
	catch(CException &e)
	{
		if(doLog)
		{
			m_logHelper.LogNK2(L"GetNK2ContactsCount Exception Catch START:");			
			TCHAR err[MAX_PATH];
			e.GetErrorMessage(err,MAX_PATH);
			m_logHelper.LogNK2(err);
			m_logHelper.LogNK2(L"GetNK2ContactsCount Exception Catch END:");
			m_logHelper.CloseNK2Log();
		}
	}
	return count;
}



BOOL CNK2Parser::ExtractNK2UDataInTextFile(TCHAR* filePath,TCHAR* error)
{
	if(!CNK2Parser::IsNK2FileExist())
	{
		wcscpy(error, L"TRContactFinder cannot find NK2 file.");
		return FALSE;
	}

	/*if(!CNK2Parser::ExtractBinResource(L"BIN", IDR_BIN_NK2VIEW_U, L"nk2viewu.exe"))
	{
		wcscpy(error, L"TRContactFinder cannot find NK2 file.");
		return FALSE;
	}	*/
	
	//tstring nk2TextFile = L"C:\\nk2u.txt";

	tstring nk2TextFile = CCommonMethods::GetLocalAppDataFolder() + L"\\nk2u.txt";

	//tstring tsCmdLine = L"C:\\nk2viewu /stab " + nk2TextFile;	
	tstring tsCmdLine = L"\"" + CCommonMethods::GetLocalAppDataFolder()  + L"\\nk2viewu.exe\" /stab \"" + nk2TextFile + L"\"";	
	
	char szCmdline[_MAX_PATH];
	memset(szCmdline,'\0',_MAX_PATH);

	wcstombs(szCmdline,tsCmdLine.c_str(),_MAX_PATH);

	UINT res = WinExec(szCmdline,0);
	switch(res)
	{
	case ERROR_BAD_FORMAT:
	case ERROR_FILE_NOT_FOUND:
	case ERROR_PATH_NOT_FOUND:

		wcscpy(error, L"TRContactFinder cannot find NK2 file.");
		return FALSE;		

	case ERROR_ACCESS_DENIED:
		wcscpy(error, L"Access Denied to NK2 file.");
		return FALSE;

	case ERROR_ACCESS_DISABLED_BY_POLICY:
		wcscpy(error, L"Accress Denied to NK2 file by policy.");
		return FALSE;
	}
			
	wcscpy(filePath,nk2TextFile.c_str());

	return TRUE;
}
BOOL CNK2Parser::IsNK2FileExist()
{
	WCHAR appDataFolder[MAX_PATH];
	if(SHGetSpecialFolderPath(NULL,appDataFolder,CSIDL_APPDATA,FALSE))
	{
		WCHAR filePath1[MAX_PATH];
		wcscpy(filePath1,appDataFolder);
		wcscat(filePath1,L"\\Microsoft\\Outlook\\outlook.nk2");

		WCHAR filePath2[MAX_PATH];
		wcscpy(filePath2,appDataFolder);
		wcscat(filePath2,L"\\Roaming\\Microsoft\\Outlook\\outlook.nk2");

		if(CCommonMethods::IsFileExist(filePath1) || CCommonMethods::IsFileExist(filePath2))
		{
			return TRUE;
		}
	}
	return FALSE;
}


BOOL CNK2Parser::ExtractBinResource( LPCWSTR strCustomResName, int nResourceId, tstring strOutputName )
{
	HGLOBAL hResourceLoaded;		// handle to loaded resource 
	HRSRC hRes;						// handle/ptr. to res. info. 
	char *lpResLock;				// pointer to resource data 
	DWORD dwSizeRes;
	tstring strOutputLocation;
	tstring strAppLocation; 		
	
	//strOutputLocation = L"C:\\"+strOutputName;
	strOutputLocation = CCommonMethods::GetLocalAppDataFolder() + L"\\" +strOutputName;
	
	CFileFind ff;
	if(ff.FindFile(strOutputLocation.c_str()))
	{
		return TRUE;
	}
	
	HMODULE hModule = AfxGetResourceHandle();

	// find location of the resource and get handle to it
	hRes = FindResource(hModule , MAKEINTRESOURCE(nResourceId), strCustomResName );
	if(hRes == NULL)
		return FALSE;
	
	// loads the specified resource into global memory. 
	hResourceLoaded = LoadResource( hModule, hRes ); 

	// get a pointer to the loaded resource!
	lpResLock = (char*)LockResource( hResourceLoaded ); 

	// determine the size of the resource, so we know how much to write out to file!  
	dwSizeRes = SizeofResource( hModule, hRes );

	std::ofstream outputFile(strOutputLocation.c_str(), std::ios::binary);

	outputFile.write((const char*)lpResLock, dwSizeRes);
	outputFile.close();

	return TRUE;
}

BOOL CNK2Parser::ConvertNK2File2TextFile(TCHAR* txtFilePath)
{	
	TCHAR nk2FilePath[MAX_PATH];
	if(GetNK2File(nk2FilePath))
	{				
		std::vector<byte> fileData;
		if(!CCommonMethods::GetFileData(tstring(nk2FilePath),fileData))	
			return FALSE;

		wcscpy(txtFilePath, CString(CCommonMethods::GetLocalAppDataFolder().c_str()) + L"\\nk2.txt");
		CTextFileIO::EncodingType eType(CTextFileIO::UTF16_LE);
		CTextFileIO txtFile(eType);
		txtFile.OpenW(txtFilePath,L"wb");

		if(!fileData.empty())
		{			
			int skipWords = 0;
			size_t size = fileData.size();
			size_t i =0;
			for(i=0;i<size;)
			{				
				if(skipWords < 5)//skip first 5 words probably file header
				{
					int skip = 1;				

					if(IsWord(&fileData.at(i),skip))				
						skipWords++;                
					
					i += skip;

					if(skipWords == 5)
						break;
		               
					continue;
				}        
			}

			std::vector<byte> rowData;    		

			for(;i<size;)
			{			
				BYTE b = (BYTE)fileData.at(i);
				
				if(IsRowEnd(&fileData.at(i)))
				{			
					tstring line = ProcessRow(rowData);

					if(!line.empty())
					{
						txtFile.WriteLineW(line);
						line.clear();
					}

					rowData.clear();

					i += 7; // skip row end seperator bytes, and point to start of next row
				}
				else
				{
					rowData.push_back(b);
					i++;
				}
			}		
		}
		txtFile.Close();		
		return TRUE;
	}
	return FALSE;
}


BOOL CNK2Parser::GetNK2File(TCHAR* nk2File)
{
	WCHAR appDataFolder[MAX_PATH];
	if(SHGetSpecialFolderPath(NULL,appDataFolder,CSIDL_APPDATA,FALSE))
	{
		WCHAR filePath1[MAX_PATH];
		wcscpy(filePath1,appDataFolder);
		wcscat(filePath1,L"\\Microsoft\\Outlook\\outlook.nk2");

		WCHAR filePath2[MAX_PATH];
		wcscpy(filePath2,appDataFolder);
		wcscat(filePath2,L"\\Roaming\\Microsoft\\Outlook\\outlook.nk2");

		if(CCommonMethods::IsFileExist(filePath1))
		{
			wcscpy(nk2File,filePath1);
			return TRUE;
		}
		else if(CCommonMethods::IsFileExist(filePath2))
		{
			wcscpy(nk2File,filePath2);
			return TRUE;
		}
		else 
		{
			WCHAR appDataFolder[MAX_PATH];
			if(SHGetSpecialFolderPath(NULL,appDataFolder,CSIDL_APPDATA,FALSE))
			{
				WCHAR nk2filePath[MAX_PATH];
				wcscpy(nk2filePath,appDataFolder);
				wcscat(nk2filePath,L"\\Microsoft\\Outlook\\*.nk2");

				WCHAR nk2filePath2[MAX_PATH];
				wcscpy(nk2filePath2,appDataFolder);
				wcscat(nk2filePath2,L"\\Roaming\\Microsoft\\Outlook\\*.nk2");

				CString nk2FileName;
				CFileFind ff;
				if(ff.FindFile(nk2filePath))
				{
					ff.FindNextFile();
					nk2FileName = CString(appDataFolder) + L"\\Microsoft\\Outlook\\" + ff.GetFileName();					
					wcscpy(nk2File,nk2FileName);
					return TRUE;
				}
				else if(ff.FindFile(nk2filePath2))
				{
					ff.FindNextFile();
					nk2FileName = CString(appDataFolder) + L"\\Roaming\\Microsoft\\Outlook\\" + ff.GetFileName();
					wcscpy(nk2File,nk2FileName);
					return TRUE;
				}
			}
		}
	}
	return FALSE;
}


tstring CNK2Parser::ProcessRow(std::vector<byte> rowData)
{	
	tstring type = GetType(rowData);
	
	if(type == L"Exchange")
	{
		return ProcessExchangeRow(rowData);
	}
	else if(type == L"SMTP" || type == L"MAPIDL")
	{
		return ProcessSMTPRow(rowData);
	}	
	return L"";
}

tstring CNK2Parser::ProcessExchangeRow(std::vector<byte> rowData)
{
	tstring line;
	tstring name;
	tstring email;
	
	std::vector<byte> nameBytes;
	std::vector<byte> emailBytes;
	
	size_t size = rowData.size();
	size_t i = 0;
    BOOL emailStart = FALSE;

	BYTE emailStartBytes[] = {0x20,0x00,0x20,0x00,0x3c,0x00};
	BYTE emailEndBytes[] = {0x00,0x3e};

	for(i=0;i<size;)
	{
		if((memcmp(emailStartBytes,&rowData.at(i),6) == 0))
		{
			i += 6;
			emailStart = TRUE;	
			break;
		}
		else
		{
			i++;
			continue;
		}		
	}
	if(emailStart)
	{
		for(;i<size;)
		{
			if((memcmp(emailEndBytes,&rowData.at(i),2) == 0))
			{
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);

				email = BytesToWideString(emailBytes);				

				//only last 3 zeros
				emailBytes.pop_back();
				emailBytes.pop_back();
				emailBytes.pop_back();
				
				break;
			}
			else			
			{
				emailBytes.push_back(rowData.at(i));
				i++;
			}
		}
		
		i = i - (emailBytes.size() + 6);
		i--;
		emailBytes.clear();

		for(;i>=0;)
		{
			BYTE b = (BYTE)rowData.at(i);

			if(!AreThreeZeros(&rowData.at(i-2)))
			{
				nameBytes.push_back(b);
				i--;
				continue;
			}
			else
			{
				i -= 2;		
				Reverse(nameBytes);
				
				nameBytes.push_back(0x00);
				nameBytes.push_back(0x00);
				nameBytes.push_back(0x00);
				nameBytes.push_back(0x00);

				name = BytesToWideString(nameBytes);
				nameBytes.clear();

				break;
			}
		}		
	}				
	
	//if not valid or empty email then dig more to find the email in the rowData
	if(!IsValidEmail(email) || email.empty())
	{						
		email = ExtractEmailFromExchangeContact(rowData);
	}

	if(!name.empty() && !email.empty())
	{
		line = name + L"\t" + email + L"\t" + L"Exchange";
	}

	return line;
}
tstring CNK2Parser::ProcessSMTPRow(std::vector<byte> rowData)
{
	tstring line;
	tstring name;
	tstring email;
	tstring type = GetType(rowData);

	std::vector<byte> nameBytes;
	std::vector<byte> emailBytes;
	
	BOOL dataStart = FALSE;		
		
	size_t size = rowData.size();
	size_t i = 0;

	BYTE COL_SEP1[] = {0x2B,0x1F,0xA4,0xBE,0xA3,0x10,0x19,0x9D,0x6E,0x00,0xDD,0x01,0x0F,0x54,0x02,0x00,0x00,0x01,0x80};
	BYTE COL_SEP2[] = {0x2B,0x1F,0xA4,0xBE,0xA3,0x10,0x19,0x9D,0x6E,0x00,0xDD,0x01,0x0F,0x54,0x02,0x00,0x00,0x01,0x90};   
  
	for(;i<size;)
	{				
		dataStart = (memcmp(COL_SEP1,&rowData.at(i),19)==0) || (memcmp(COL_SEP2,&rowData.at(i),19) == 0);

        if(!dataStart)
		{
			i++;
			continue;
		}		
		else
		{
			i += 19;  //skip seperator bytes
			break;;
		}															
	}
	
	if(dataStart)
	{
        std::vector<byte> fieldBytes;
		BOOL wordStart = TRUE;		
        int wordCount = 0;

		for(;i<size;)
		{			
			BYTE b = (BYTE)rowData.at(i);

            if(!AreThreeZeros(&rowData.at(i)))
			{
				fieldBytes.push_back(b);
				i++;
				continue;
			}
			else
			{
				i += 3; //skip 3 NULL bytes
				wordCount++;

                //put 4 zero, 1 for end of field and 3 for break among fields
				fieldBytes.push_back(0x00);
				fieldBytes.push_back(0x00);
				fieldBytes.push_back(0x00);
                fieldBytes.push_back(0x00);
				
				if(wordCount == 1) //email
				{											
					name = BytesToWideString(fieldBytes);
				}
				//if(wordCount == 2) //type
				//{											
				//	type = BytesToWideString(fieldBytes);
				//}
				if(wordCount == 3) //name
				{						
					email = BytesToWideString(fieldBytes);
				}
				// reset
				fieldBytes.clear();
			}

			if(wordCount == 3)
				break;;			
		}
	}

	if(!dataStart || email.empty())
	{
		i = rowData.size()-1;
		BOOL emailStart = FALSE;

		BOOL rowHasNameAndEmail = FALSE;
		if(size > 0)
		{
			if(rowData.at(size - 1) == 0x3e)
				rowHasNameAndEmail = TRUE;
		}
		if(rowHasNameAndEmail)
		{
			BYTE emailStartBytes[] = {0x20,0x00,0x20,0x00,0x3c,0x00};
			BYTE emailEndBytes[] = {0x00,0x3e};

			size = rowData.size();

			i = 0;
			emailStart = FALSE;

			for(i=0;i<size;)
			{
				if((memcmp(emailStartBytes,&rowData.at(i),6) == 0))
				{
					i += 6;
					emailStart = TRUE;	
					break;
				}
				else
				{
					i++;
					continue;
				}		
			}
			if(emailStart)
			{
				for(;i<size;)
				{
					if((memcmp(emailEndBytes,&rowData.at(i),2) == 0))
					{
						emailBytes.push_back(0x00);
						emailBytes.push_back(0x00);
						emailBytes.push_back(0x00);
						emailBytes.push_back(0x00);

						email = BytesToWideString(emailBytes);				

						//only last 3 zeros
						emailBytes.pop_back();
						emailBytes.pop_back();
						emailBytes.pop_back();
						
						break;
					}
					else			
					{
						emailBytes.push_back(rowData.at(i));
						i++;
					}
				}
				
				i = i - (emailBytes.size() + 6);
				i--;
				emailBytes.clear();

				for(;i>=0;)
				{
					BYTE b = (BYTE)rowData.at(i);

					if(!AreThreeZeros(&rowData.at(i-2)))
					{
						nameBytes.push_back(b);
						i--;
						continue;
					}
					else
					{
						i -= 2;		
						Reverse(nameBytes);
				
						nameBytes.push_back(0x00);
						nameBytes.push_back(0x00);
						nameBytes.push_back(0x00);
						nameBytes.push_back(0x00);

						name = BytesToWideString(nameBytes);
						
						nameBytes.clear();

						break;
					}
				}		
			}				
		}
		else
		{
			for(;i>=0;)
			{
				BYTE b = (BYTE)rowData.at(i);

				if(!AreThreeZeros(&rowData.at(i-2)))
				{
					nameBytes.push_back(b);
					i--;
					continue;
				}
				else
				{
					i -= 2;		
					Reverse(nameBytes);
				
					nameBytes.push_back(0x00);
					nameBytes.push_back(0x00);
					nameBytes.push_back(0x00);
					nameBytes.push_back(0x00);

					name = BytesToWideString(nameBytes);
					email = name;
					
					nameBytes.clear();

					break;
				}
			}	
		}				
	}
	if(!IsValidEmail(email))
	{						
		email = L"";
	}
	if(!name.empty() && !email.empty())
	{
		line = name + L"\t" + email + L"\t" + type;
	}

	return line;

}
BOOL CNK2Parser::IsRowEnd(BYTE* sep)
{
	//row seperator
	BYTE ROW_SEP[] = {0x00,0x00,0x00,0x03,0x00,0x04,0x60};
	int r = memcmp(ROW_SEP,sep,7);
	return (r==0);
}


BOOL CNK2Parser::IsWord(BYTE* sep,int& skip)
{
	//column seperator
	BYTE MAX_NULLS[] = {0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00};

	if(memcmp(MAX_NULLS,sep,12) == 0)
	{
		skip = 12;
		return FALSE;
	}
	if(memcmp(MAX_NULLS,sep,9) == 0)
	{
		skip = 9;
		return FALSE;
	}
	if(memcmp(MAX_NULLS,sep,7) == 0)
	{
		skip = 7;
		return FALSE;
	}
	if(memcmp(MAX_NULLS,sep,5) == 0)
	{
		skip = 5;
		return FALSE;
	}
	if(memcmp(MAX_NULLS,sep,4) == 0)
	{
		skip = 4;
		return FALSE;
	}		
	if(memcmp(MAX_NULLS,sep,3) == 0)
	{
		skip = 3;
		return TRUE;
	}
	if(memcmp(MAX_NULLS,sep,2) == 0)
	{
		skip = 2;
		return TRUE;
	}
	skip = 1;
	return FALSE;
}

BOOL CNK2Parser::AreThreeZeros(BYTE* sep)
{
	BYTE THREE_NULLS[] = {0x00,0x00,0x00};
	return (memcmp(THREE_NULLS,sep,3) == 0);
}
tstring CNK2Parser::BytesToWideString(std::vector<byte> bytes)
{	
	tstring wideString;	
	size_t size = bytes.size();

	for(size_t i=0;i<size-4;) // last 4 bytes are field seperator
	{									
        if(AreThreeZeros(&bytes.at(i)))
		    break;
		
		BYTE b1 = bytes.at(i++);
		BYTE b2 = bytes.at(i++);			
		wideString += (wchar_t)MAKEWORD(b1,b2); //concatenate 2 bytes in in 1 wchar
	}
	return wideString;
}

void CNK2Parser::Reverse(std::vector<byte>& bytes)
{
	std::vector<byte> tempBytes = bytes;
	bytes.clear();
	int size = tempBytes.size();
	for(int i=size-1;i>=0;i--)
	{
		bytes.push_back(tempBytes.at(i));
	}
}
BOOL CNK2Parser::IsValidEmail(const tstring email)
{
	return (email.find(L"@") != std::string::npos) && (email.size() >= 5);
}
tstring CNK2Parser::GetType(std::vector<byte> rowBytes)
{
	BYTE EX_BYTES[] = {0x00,0x00,0x00,0x45,0x00,0x58,0x00,0x00,0x00};//...E.X...
	
    for(size_t i=0;i<rowBytes.size();i++)
    {
        if(i+9 < rowBytes.size())
        {
            if((memcmp(EX_BYTES,&rowBytes.at(i),9) == 0))
                return L"Exchange";
        }
		
		else
			break;
    }
	
	//SMTP
	BYTE SMTP_BYTES[] = {0x00,0x00,0x00,0x53,0x00,0x4d,0x00,0x54,0x00,0x50,0x00,0x00,0x00};//...S.M.T.P...

	for(size_t i=0;i<rowBytes.size();i++)
    {
        if(i+13 < rowBytes.size())
        {
            if((memcmp(SMTP_BYTES,&rowBytes.at(i),13) == 0))
                return L"SMTP";
        }
		else
			break;
    }

	//distribution list
	BYTE MAPIDL_BYTES[] = {0x00,0x00,0x00,0x4d,0x00,0x41,0x00,0x50,0x00,0x49,0x00,0x50,0x00,0x44,0x00,0x4c,0x00,0x00,0x00};//...M.A.P.I.D.L...

	for(size_t i=0;i<rowBytes.size();i++)
    {
        if(i+19 < rowBytes.size())
        {
            if((memcmp(MAPIDL_BYTES,&rowBytes.at(i),19) == 0))
                return L"MAPIDL";
        }
		else
			break;
    }	

	return L"";


}
tstring CNK2Parser::ExtractEmailFromExchangeContact(std::vector<byte> rowBytes)
{
	tstring email = L" ";
	std::vector<byte> emailBytes;
	size_t i = 0;
	size_t size = rowBytes.size();
	BOOL wordStart = FALSE;
	int skip = 1;

	for(i=0;i<size;)
	{
		BYTE b = rowBytes.at(i);
		if(IsWord(&rowBytes.at(i),skip))
		{
			wordStart = !wordStart;
			
			if(!wordStart) //End Word
			{
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);
				emailBytes.push_back(0x00);
				
				email = BytesToWideString(emailBytes);				

				emailBytes.clear();	

				if(IsValidEmail(email))
					return email;
				else
					return L"";
			}
			i += skip;						
		}
		else
		{
			if(wordStart)
				emailBytes.push_back(b);
			i++;
			continue;
		}		
	}
	
	return email;
}
