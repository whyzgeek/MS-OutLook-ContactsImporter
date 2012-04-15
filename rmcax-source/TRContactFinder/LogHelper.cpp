#include "StdAfx.h"
#include "LogHelper.h"
#include "TextFileIO.h"
#include "CommonMethods.h"

CLogHelper::CLogHelper(void)
{			
}

CLogHelper::~CLogHelper(void)
{
	if(m_PABLogFile.IsOpen())
		m_PABLogFile.Close();
	if(m_NK2LogFile.IsOpen())
		m_NK2LogFile.Close();
}
BOOL CLogHelper::OpenPABLogFile()
{
	if(m_PABLogFile.IsOpen())
		return TRUE;

	TCHAR tempPath[_MAX_PATH];
	memset(tempPath,'\0',_MAX_PATH);
	GetTempPath(_MAX_PATH,tempPath);	

	if(CCommonMethods::IsDirExist(tempPath,L"TrContactFinderLog"))
	{
		TCHAR fileName[_MAX_PATH];
		wcscpy(fileName,tempPath);
		wcscat(fileName,L"TrContactFinderLog\\");
		wcscat(fileName,L"PABLog.txt");

		return m_PABLogFile.OpenW(fileName,L"w");
	}
	return FALSE;
}

BOOL CLogHelper::OpenNK2LogFile()
{
	if(m_NK2LogFile.IsOpen())
		return TRUE;

	TCHAR tempPath[_MAX_PATH];
	memset(tempPath,'\0',_MAX_PATH);
	GetTempPath(_MAX_PATH,tempPath);
	
	if(CCommonMethods::IsDirExist(tempPath,L"TrContactFinderLog"))
	{
		TCHAR fileName[_MAX_PATH];
		wcscpy(fileName,tempPath);
		wcscat(fileName,L"TrContactFinderLog\\");
		wcscat(fileName,L"NK2Log.txt");

		return m_NK2LogFile.OpenW(fileName,L"w");
	}
	return FALSE;
}

void CLogHelper::LogPAB(const wchar_t* text)
{
	if(OpenPABLogFile())
	{
		m_PABLogFile.WriteLine(text);
	}
}

void CLogHelper::LogNK2(const wchar_t* text)
{
	if(OpenNK2LogFile())
	{
		m_NK2LogFile.WriteLine(text);		
	}	
}

void CLogHelper::ClosePABLog()
{
	m_PABLogFile.Close();
}
void CLogHelper::CloseNK2Log()
{
	m_NK2LogFile.Close();
}
