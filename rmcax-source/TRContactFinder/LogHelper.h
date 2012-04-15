#pragma once
#include "TextFileIO.h"

class CLogHelper
{
private:

	CTextFileIO m_PABLogFile;
	CTextFileIO m_NK2LogFile;
	BOOL OpenPABLogFile();
	BOOL OpenNK2LogFile();

public:
	CLogHelper(void);
	~CLogHelper(void);

	void LogPAB(const wchar_t* text);
	void LogNK2(const wchar_t* text);
	void ClosePABLog();
	void CloseNK2Log();
};
