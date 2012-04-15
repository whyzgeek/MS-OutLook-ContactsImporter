////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: TextFileIO.cpp
// Description: Implementation of CTextFileIO
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "io.h"
#include "TextFileIO.h"

CTextFileIO::CTextFileIO(void)
:encodingType(ANSI)
, _IsValidate(false)
{
	_file = NULL;
}

CTextFileIO::CTextFileIO(EncodingType eType)
:_IsValidate(false)
{
	encodingType = eType;
	_file = NULL;
}

CTextFileIO::~CTextFileIO(void)
{
	if(_file)
	{
		fclose(_file);
		_file = NULL;
	}
}

CTextFileIO::CTextFileIO(const WCHAR* filename,  WCHAR *mode)
{
	// Testing for file type
	encodingType=CTextFileIO::CheckFileEncodingTypeW(filename);

	_IsValidate=OpenW(filename,mode);	
}

CTextFileIO::CTextFileIO(const char* filename,  char *mode)
{
	// Testing for file type
	encodingType=CTextFileIO::CheckFileEncodingTypeA(filename);

	_IsValidate=OpenA(filename,mode);	
}

// Open file,UNICODE version
BOOL CTextFileIO::OpenW(const WCHAR* const filename, WCHAR *mode)
{
	//_file=_wfopen_s((filename,mode);	
	errno_t err  =_wfopen_s(&_file,filename,mode);	
	if(_file==NULL)
		return FALSE;
	// Write Encoding tag
	if(wcschr(mode,L'w')!=NULL)
		WriteEncodingTag(encodingType);
	// Seek  file pos
	if(wcschr(mode,L'r')!=NULL)
		OmitEncodingTag(encodingType);
	
	return TRUE;
}

// Open file,ANSI Version
BOOL CTextFileIO::OpenA(const char* const filename, char *mode)
{
	//_file=fopen(filename,mode);
	errno_t err  =fopen_s(&_file,filename,mode);	

	if(_file==NULL)
		return FALSE;
	// Write Encoding tag
	if(strchr(mode,'w')!=NULL)
		WriteEncodingTag(encodingType);
	// Seek  file pos
	if(strchr(mode,'r')!=NULL)
		OmitEncodingTag(encodingType);
	
	return TRUE;
}
BOOL CTextFileIO::Close()
{
	int close = -1;
	if(_file)
	{
		close = fclose(_file);
		if(close == 0)
			_file = NULL;
	}
	return close == 0 ? TRUE : FALSE;
}
bool CTextFileIO::_ReadLine(string& s, int eol, int eof)
{

    // reset string
    s.clear();
    
    // read one char at a time
    while (true)
    {
        // read char
        int c = fgetc(_file);        
        
        // check for EOF
        if (c == eof || c == EOF) return false;

        // check for EOL
        if (c == eol) return true;

        // append this character to the string
        s += c;
    };
}

// Read a line from file,return value is UNICODE string
bool CTextFileIO::_ReadLine(wstring& s, wint_t eol, wint_t eof)
{
    // reset string
    s.clear();

    
    // read one wide char at a time
    while (true)
    {
        // read wide char
        wint_t c = fgetwc(_file);
        
        // check for EOF
        if (c == eof || c == WEOF) return false;

        // check for EOL
        if (c == eol) return true;

        // append the wide character to the string        
        s += c;
    };
	return 0;
}

bool CTextFileIO::_WriteLine(const char *const s, int ret,int newline, size_t length)
{
	// check if the pointer is valid
    if (!s)
    {
        return false;
    };
    
    // calculate the string's length
    if (length==-1)
    {
        length = strlen(s);
    };
    
    // write the string to the file
    size_t n = fwrite(s, sizeof(char), length, _file);
            
    // write line break to the file
	fputc(ret, _file);
	fputc(newline, _file);
            
    // return whether the write operation was successful
    return (n == length);
}

bool CTextFileIO::_WriteLine(const wchar_t *const s, wint_t ret,wint_t newline, size_t length)
{
	    // check if the pointer is valid
    if (!s)
    {
        return false;
    };
        
    // calculate the string's length
    if (length==-1)
    {
        length = wcslen(s);
    };    
    
    // write the string to the file
    size_t n = fwrite(s, sizeof(wchar_t), length, _file);
            
    // write line break to the file
	fputwc(ret, _file);
	fputwc(newline,_file);
	/*BYTE retBytes[]={0x0d,0x00};
	fwrite(retBytes,sizeof(BYTE),2,_file);

	BYTE newLineBytes[]={0x0a,0x00};
	fwrite(newLineBytes,sizeof(BYTE),2,_file);*/
            
    // return whether the write operation was successful
    return (n == length);
}

// Unicode version of read line
bool CTextFileIO::ReadLineW(std::wstring &ws)
{
	bool bResult=false;
	switch(encodingType)
	{
	case ANSI:
		{
			std::string s;
			bResult=_ReadLine(s);
			s+='\0';
			int nLength=MultiByteToWideChar(CP_OEMCP,MB_PRECOMPOSED,s.c_str(),-1,NULL,0);
			LPWSTR lpwStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_OEMCP,MB_PRECOMPOSED,s.c_str(),s.length(),lpwStr,nLength);
			ws=lpwStr;
		}
		break;
	case UTF_8:
		{
			std::string s;
			bResult=_ReadLine(s);
			s+='\0';
			int nLength=MultiByteToWideChar(CP_UTF8,0,s.c_str(),-1,NULL,0);
			LPWSTR lpwStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_UTF8,0,s.c_str(),-1,lpwStr,nLength);
			ws=lpwStr;
		}
		break;
	case UTF16_BE:
		{
			std::wstring s;
			bResult=_ReadLine(s,0x0A00);
			// Convert UTF16 big endian to UTF little endian
			int nLength=s.length()*2-2;
			BYTE *src=(BYTE*)s.c_str();
			BYTE *dst=new BYTE[nLength+2];
			memset(dst,0,nLength+2);
			for(int i=0;i<nLength;)
			{
				dst[i]=src[i+1];
				dst[i+1]=src[i];
				i+=2;
			}
			ws=(WCHAR*)dst;
		}
		break;
	case UTF16_LE:
		{
			std::wstring s;
			bResult=_ReadLine(s);
			ws=s;
		}
		break;
	case UTF32_BE:
		{
			// TODO:will be implement later
		}
		break;
	case UTF32_LE:
		{
			// TODO:will be implement later
		}
		break;
	default:
		{
		}
	}
	return bResult;
}

bool CTextFileIO::ReadLineW(LPCWSTR ws)
{
	std::wstring s;
	bool bResult=ReadLineW(s);
	ws=s.c_str();	
	return bResult;
}
// ANSI version of read line
bool CTextFileIO::ReadLineA(std::string &s)
{
	bool bResult=false;
	switch(encodingType)
	{
	case ANSI:
		{
			bResult=_ReadLine(s);
		}
		break;
	case UTF_8:
		{
			std::string s;
			bResult=_ReadLine(s);
			s+='\0';
			// Convert utf-8 to ANSI,must first convert to UTF16-LE,then to ANSI
			int nLength=MultiByteToWideChar(CP_UTF8,0,s.c_str(),-1,NULL,0);
			LPWSTR lpwStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_UTF8,0,s.c_str(),-1,lpwStr,nLength);
			nLength=WideCharToMultiByte(CP_ACP,0,lpwStr,nLength,NULL,0,NULL,NULL);
			LPSTR lpStr=new CHAR[nLength];
			WideCharToMultiByte(CP_UTF8,0,lpwStr,nLength,lpStr,nLength,NULL,NULL);
			s=lpStr;
		}
		break;
	case UTF16_BE:
		{
			// Convert UTF16 big endian to UTF little endian
			std::wstring ws;
			bResult=_ReadLine(ws,0x0A00);
			int nLength=ws.length()*2-2;
			BYTE *src=(BYTE*)ws.c_str();
			BYTE *dst=new BYTE[nLength+2];
			memset(dst,0,nLength+2);
			for(int i=0;i<nLength;)
			{
				dst[i]=src[i+1];
				dst[i+1]=src[i];
				i+=2;
			}
			ws.clear();
			ws=(WCHAR*)dst;
			// Convert UTF16 little endian to ANSI
			nLength=WideCharToMultiByte(CP_ACP,0,ws.c_str(),-1,NULL,0,NULL,NULL);
			LPSTR lpStr=new CHAR[nLength];
			WideCharToMultiByte(CP_ACP,0,ws.c_str(),-1,lpStr,nLength,NULL,NULL);
			s=lpStr;
		}
		break;
	case UTF16_LE:
		{
			std::wstring ws;
			bResult=_ReadLine(ws);
			int nLength=WideCharToMultiByte(CP_ACP,0,ws.c_str(),-1,NULL,0,NULL,NULL);
			LPSTR lpStr=new CHAR[nLength];
			WideCharToMultiByte(CP_ACP,0,ws.c_str(),-1,lpStr,nLength,NULL,NULL);
			s=lpStr;
		}
		break;
	case UTF32_BE:
		{
			// TODO:will be implement later			
		}
		break;
	case UTF32_LE:
		{
			// TODO:will be implement later
		}
		break;
	default:
		{
		}
	}
	return bResult;
}

bool CTextFileIO::ReadLineA(LPCSTR s)
{
	std::string s_s;
	bool bResult=ReadLineA(s_s);
	s=s_s.c_str();	
	return bResult;
}

// UNICODE version of write line
bool CTextFileIO::WriteLineW(const wchar_t *const wc)
{
	bool bResult=false;
	switch(encodingType)
	{
	case ANSI:
		{
			int nLength=WideCharToMultiByte(CP_ACP,0,wc,-1,NULL,0,NULL,NULL);
			LPSTR lpStr=new char[nLength];
			WideCharToMultiByte(CP_ACP,0,wc,-1,lpStr,nLength,NULL,NULL);
			bResult=_WriteLine(lpStr);
		}
		break;
	case UTF_8:
		{
			// Convert utf-16 to UTF8
			int nLength=WideCharToMultiByte(CP_UTF8,0,wc,-1,NULL,0,NULL,NULL);
			LPSTR lpStr=new CHAR[nLength];
			WideCharToMultiByte(CP_UTF8,0,wc,-1,lpStr,nLength,NULL,NULL);
			// Write to file
			bResult=_WriteLine(lpStr);
		}
		break;
	case UTF16_BE:
		{
			BYTE *src=(BYTE*)wc;
			int nLength=wcslen(wc)*2;
			BYTE *dst=new BYTE[nLength]+2;
			memset(dst,0,nLength+2);
			for(int i=0;i<nLength;)
			{
				dst[i]=src[i+1];
				dst[i+1]=src[i];
				i+=2;
			}
			bResult=_WriteLine((wchar_t*)dst,0x0D00,0x0A00);
		}
		break;
	case UTF16_LE:
		{
			bResult=_WriteLine(wc);
		}
		break;
	case UTF32_BE:
		{
			// TODO:will be implement later
		}
		break;
	case UTF32_LE:
		{
			// TODO:will be implement later
		}
		break;
	default:
		{
		}
	}
	return bResult;
}

// ANSI Version of writeline
bool CTextFileIO::WriteLineA(const char *const c)
{
	bool bResult=false;
	switch(encodingType)
	{
	case ANSI:
		{
			bResult=_WriteLine(c);
		}
		break;
	case UTF_8:
		{
			// First convert to UTF16 litter endier
			int nLength=MultiByteToWideChar(CP_ACP,0,c,-1,NULL,0);
			LPWSTR lpWStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_ACP,0,c,-1,lpWStr,nLength);
			// Convert utf-16 to UTF8
			int Length=WideCharToMultiByte(CP_UTF8,0,lpWStr,nLength,NULL,0,NULL,NULL);
			LPSTR lpStr=new CHAR[Length];
			WideCharToMultiByte(CP_UTF8,0,lpWStr,nLength,lpStr,Length,NULL,NULL);
			// Write to file
			bResult=_WriteLine(lpStr);
		}
		break;
	case UTF16_BE:
		{
			// First convert to UTF16 litter endian
			int nLength=MultiByteToWideChar(CP_ACP,0,c,-1,NULL,0);
			LPWSTR lpWStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_ACP,0,c,-1,lpWStr,nLength);
			// Then convert to UTF16 big endian
			BYTE *src=(BYTE*)lpWStr;
			nLength=wcslen(lpWStr)*2;
			BYTE *dst=new BYTE[nLength+2];
			memset(dst,0,nLength+2);
			for(int i=0;i<nLength;)
			{
				dst[i]=src[i+1];
				dst[i+1]=src[i];
				i+=2;
			}
			// Write to file
			bResult=_WriteLine((wchar_t*)dst,0x0D00,0x0A00);
		}
		break;
	case UTF16_LE:
		{
			// First convert to UTF16 litter endian
			int nLength=MultiByteToWideChar(CP_ACP,0,c,-1,NULL,0);
			LPWSTR lpWStr=new WCHAR[nLength];
			MultiByteToWideChar(CP_ACP,0,c,-1,lpWStr,nLength);
			bResult=_WriteLine(lpWStr);
		}
		break;
	case UTF32_BE:
		{
			// TODO:will be implement later
		}
		break;
	case UTF32_LE:
		{
			// TODO:will be implement later
		}
		break;
	default:
		{
		}
	}
	return bResult;
}
// Check file encoding type
CTextFileIO::EncodingType CTextFileIO::CheckFileEncodingTypeW(const WCHAR* const filename)
{
	//FILE* _file=_wfopen(filename,L"rb");
	FILE* _file;
	errno_t err  =_wfopen_s(&_file,filename,L"rb");	
	// Read first 4 byte for testing
	BYTE* buf=new BYTE[4];
	size_t nRead=fread((void*)buf,sizeof(BYTE),4,_file);
	// Close file
	fclose(_file);

	// Testing
	if(nRead<2)
		return ANSI;
	// Testting tocken
	BYTE utf32_le[]={0xFF,0xFE,0x00,0x00};
	if(memcmp(buf,&utf32_le,4)==0)
		return UTF32_LE;

	BYTE utf32_be[]={0x00,0x00,0xFE,0xFF};
	if(memcmp(buf,&utf32_be,4)==0)
		return UTF32_BE;

	BYTE utf_8[]={0xEF,0xBB,0xBF};
	if(memcmp(buf,&utf_8,3)==0)
		return UTF_8;

	BYTE utf16_le[]={0xFF,0xFE};
	if(memcmp(buf,&utf16_le,2)==0)
		return UTF16_LE;

	BYTE utf16_be[]={0xFE,0xFF};
	if(memcmp(buf,&utf16_be,2)==0)
		return UTF16_BE;
	// Else
	return ANSI;
}

CTextFileIO::EncodingType CTextFileIO::CheckFileEncodingTypeA(const CHAR* const filename)
{
	//FILE* _file=fopen(filename,"rb");
	FILE* _file;
	errno_t err  =fopen_s(&_file,filename,"rb");	
	// Read first 4 byte for testing
	BYTE* buf=new BYTE[4];
	size_t nRead=fread((void*)buf,sizeof(BYTE),4,_file);
	// Close file
	fclose(_file);

	// Testing
	if(nRead<2)
		return ANSI;
	// Testting tocken
	BYTE utf32_le[]={0xFF,0xFE,0x00,0x00};
	if(memcmp(buf,&utf32_le,4)==0)
		return UTF32_LE;

	BYTE utf32_be[]={0x00,0x00,0xFE,0xFF};
	if(memcmp(buf,&utf32_be,4)==0)
		return UTF32_BE;

	BYTE utf_8[]={0xEF,0xBB,0xBF};
	if(memcmp(buf,&utf_8,3)==0)
		return UTF_8;

	BYTE utf16_le[]={0xFF,0xFE};
	if(memcmp(buf,&utf16_le,2)==0)
		return UTF16_LE;

	BYTE utf16_be[]={0xFE,0xFF};
	if(memcmp(buf,&utf16_be,2)==0)
		return UTF16_BE;
	// Else
	return ANSI;
}
// Omit file encoding tag
int CTextFileIO::OmitEncodingTag(EncodingType type)
{
	int nResult=0;
	switch(encodingType)
	{
	case UTF_8:
			nResult=fseek(_file,3,SEEK_SET);
		break;
	case UTF16_BE:
	case UTF16_LE:
			nResult=fseek(_file,2,SEEK_SET);
		break;
	case UTF32_BE:
	case UTF32_LE:
		nResult=fseek(_file,4,SEEK_SET);
		break;
	}
	return nResult;
}

// Write the encoding type tag an beginner of file
void CTextFileIO::WriteEncodingTag(EncodingType type)
{
	switch(type)
	{
	case UTF_8:
		{
			BYTE utf_8[]={0xEF,0xBB,0xBF};
			fwrite(utf_8,sizeof(BYTE),3,_file);
		}			
		break;
	case UTF16_BE:
		{
			BYTE utf16_be[]={0xFE,0xFF};
			fwrite(utf16_be,sizeof(BYTE),2,_file);
		}
		break;
	case UTF16_LE:
		{
			BYTE utf16_le[]={0xFF,0xFE};
			fwrite(utf16_le,sizeof(BYTE),2,_file);
		}			
		break;
	case UTF32_BE:
		{
			BYTE utf32_be[]={0x00,0x00,0xFE,0xFF};
			fwrite(utf32_be,sizeof(BYTE),4,_file);
		}
	case UTF32_LE:
		{
			BYTE utf32_le[]={0xFF,0xFE,0x00,0x00};
			fwrite(utf32_le,sizeof(BYTE),4,_file);
		}		
		break;
	default:; // ANSI, do nothing
	}

}


void CTextFileIO::WriteByte(BYTE* pByte)
{
	fwrite(pByte,sizeof(BYTE),1,_file);
}