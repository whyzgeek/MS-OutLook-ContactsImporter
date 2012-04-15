#pragma once

/////////////////////////////////////////////////////////////////////////////
//
// File: RegistryHelper.h
// Description: Registry key functions wrapper class
/////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////
class CRegistryHelper
/////////////////////////////////////////////////////////////////////////////
{
public:
	CRegistryHelper();
	~CRegistryHelper();

// Attributes
public:
	operator HKEY() const;
	HKEY m_hKey;

// Operations
public:
	LONG SetValue(DWORD dwValue, LPCTSTR lpszValueName);
	LONG SetValue(LPCTSTR lpszValue, LPCTSTR lpszValueName = NULL, BOOL fExpandStr = FALSE);

	LONG QueryValue(DWORD& dwValue, LPCTSTR lpszValueName);
	LONG QueryValue(CString& strValue, LPCTSTR lpszValueName, BOOL bExpandStrings = FALSE);
	LONG QueryValue(BYTE* pData, LPCTSTR lpszValueName, LPDWORD lpdwSize);

	LONG SetKeyValue(LPCTSTR lpszKeyName, LPCTSTR lpszValue, LPCTSTR lpszValueName = NULL);
	static LONG WINAPI SetValue(HKEY hKeyParent, LPCTSTR lpszKeyName,
		LPCTSTR lpszValue, LPCTSTR lpszValueName = NULL);

	LONG Create(HKEY hKeyParent, LPCTSTR lpszKeyName,
		LPTSTR lpszClass = REG_NONE, DWORD dwOptions = REG_OPTION_NON_VOLATILE,
		REGSAM samDesired = KEY_ALL_ACCESS,
		LPSECURITY_ATTRIBUTES lpSecAttr = NULL,
		LPDWORD lpdwDisposition = NULL);
	LONG Open(HKEY hKeyParent, LPCTSTR lpszKeyName,
		REGSAM samDesired = KEY_ALL_ACCESS);
	LONG Close();
	HKEY Detach();
	void Attach(HKEY hKey);
	LONG DeleteSubKey(LPCTSTR lpszSubKey);
	LONG RecurseDeleteKey(LPCTSTR lpszKey);
	LONG DeleteValue(LPCTSTR lpszValue);
	LONG RegEnumStringValues(CStringArray& strarrayValueName, 
							 CStringArray* pstrarrayValues = NULL);
	LONG GetCountChildren(DWORD& dwSubKeys, DWORD& dwValues);
	BOOL IsOpen();
};

/////////////////////////////////////////////////////////////////////////////
inline CRegistryHelper::CRegistryHelper()
/////////////////////////////////////////////////////////////////////////////
{
	m_hKey = NULL;
}

/////////////////////////////////////////////////////////////////////////////
inline CRegistryHelper::~CRegistryHelper()
/////////////////////////////////////////////////////////////////////////////
{
	Close();
}

/////////////////////////////////////////////////////////////////////////////
inline CRegistryHelper::operator HKEY() const
/////////////////////////////////////////////////////////////////////////////
{
	return m_hKey;
}

/////////////////////////////////////////////////////////////////////////////
inline LONG CRegistryHelper::SetValue(DWORD dwValue, LPCTSTR lpszValueName)
/////////////////////////////////////////////////////////////////////////////
{
	_ASSERTE(m_hKey != NULL);
	return RegSetValueEx(m_hKey, lpszValueName, NULL, REG_DWORD,
		(BYTE * const)&dwValue, sizeof(DWORD));
}

/////////////////////////////////////////////////////////////////////////////
inline HRESULT CRegistryHelper::SetValue(LPCTSTR lpszValue, LPCTSTR lpszValueName, BOOL fExpandStr)
/////////////////////////////////////////////////////////////////////////////
{
	_ASSERTE(lpszValue != NULL);
	_ASSERTE(m_hKey != NULL);
	return RegSetValueEx(m_hKey, lpszValueName, NULL, 
			fExpandStr ? REG_EXPAND_SZ : REG_SZ,
			(BYTE * const)lpszValue, (lstrlen(lpszValue)+1)*sizeof(TCHAR));
}

/////////////////////////////////////////////////////////////////////////////
inline HKEY CRegistryHelper::Detach()
/////////////////////////////////////////////////////////////////////////////
{
	HKEY hKey = m_hKey;
	m_hKey = NULL;
	return hKey;
}

/////////////////////////////////////////////////////////////////////////////
inline void CRegistryHelper::Attach(HKEY hKey)
/////////////////////////////////////////////////////////////////////////////
{
	_ASSERTE(m_hKey == NULL);
	m_hKey = hKey;
}

/////////////////////////////////////////////////////////////////////////////
inline LONG CRegistryHelper::DeleteSubKey(LPCTSTR lpszSubKey)
/////////////////////////////////////////////////////////////////////////////
{
	_ASSERTE(m_hKey != NULL);
	return RegDeleteKey(m_hKey, lpszSubKey);
}

/////////////////////////////////////////////////////////////////////////////
inline LONG CRegistryHelper::DeleteValue(LPCTSTR lpszValue)
/////////////////////////////////////////////////////////////////////////////
{
	_ASSERTE(m_hKey != NULL);
	return RegDeleteValue(m_hKey, (LPTSTR)lpszValue);
}

/////////////////////////////////////////////////////////////////////////////
inline BOOL CRegistryHelper::IsOpen()
/////////////////////////////////////////////////////////////////////////////
{
	return (m_hKey != NULL);
}
