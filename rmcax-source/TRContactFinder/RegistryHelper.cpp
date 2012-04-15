/////////////////////////////////////////////////////////////////////////////
//
// File: RegistryHelper.cpp 
// Description: Registry key functions wrapper
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "RegistryHelper.h"

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::Close()
/////////////////////////////////////////////////////////////////////////////
{
	LONG lRes = ERROR_SUCCESS;
	if (m_hKey != NULL)
	{
		lRes = RegCloseKey(m_hKey);
		m_hKey = NULL;
	}
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::Create(HKEY hKeyParent, LPCTSTR lpszKeyName,
	LPTSTR lpszClass, DWORD dwOptions, REGSAM samDesired,
	LPSECURITY_ATTRIBUTES lpSecAttr, LPDWORD lpdwDisposition)
/////////////////////////////////////////////////////////////////////////////
{
	ASSERT(hKeyParent != NULL);
	DWORD dw;
	HKEY hKey = NULL;
	LONG lRes = RegCreateKeyEx(hKeyParent, lpszKeyName, 0,
		lpszClass, dwOptions, samDesired, lpSecAttr, &hKey, &dw);
	if (lpdwDisposition != NULL)
		*lpdwDisposition = dw;
	if (lRes == ERROR_SUCCESS)
	{
		lRes = Close();
		m_hKey = hKey;
	}
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::Open(HKEY hKeyParent, LPCTSTR lpszKeyName, REGSAM samDesired)
/////////////////////////////////////////////////////////////////////////////
{
	ASSERT(hKeyParent != NULL);
	HKEY hKey = NULL;
	LONG lRes = RegOpenKeyEx(hKeyParent, lpszKeyName, 0, samDesired, &hKey);
	if (lRes == ERROR_SUCCESS)
	{
		lRes = Close();
		ASSERT(lRes == ERROR_SUCCESS);
		m_hKey = hKey;
	}
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::QueryValue(DWORD& dwValue, LPCTSTR lpszValueName)
/////////////////////////////////////////////////////////////////////////////
{
	DWORD dwType = NULL;
	DWORD dwCount = sizeof(DWORD);
	LONG lRes = RegQueryValueEx(m_hKey, (LPTSTR)lpszValueName, NULL, &dwType,
		(LPBYTE)&dwValue, &dwCount);
	ASSERT((lRes!=ERROR_SUCCESS) || (dwType == REG_DWORD));
	ASSERT((lRes!=ERROR_SUCCESS) || (dwCount == sizeof(DWORD)));
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::QueryValue(CString& strValue, LPCTSTR lpszValueName, BOOL bExpandStrings)
/////////////////////////////////////////////////////////////////////////////
{
	DWORD dwType = NULL;
	DWORD dwCount = 0;
	LONG lRes = RegQueryValueEx(m_hKey, (LPTSTR)lpszValueName, NULL, &dwType,
								NULL, &dwCount);
	
	ASSERT(lRes == ERROR_SUCCESS);
	
	ASSERT((dwType == REG_SZ) || (dwType == REG_EXPAND_SZ) || (dwType == REG_MULTI_SZ));

	// Get the string
	LPTSTR lpszStr = strValue.GetBufferSetLength(dwCount + 1);
	lRes = RegQueryValueEx(m_hKey, (LPTSTR)lpszValueName, NULL, &dwType,
								(LPBYTE)lpszStr, &dwCount);

	strValue.ReleaseBuffer();

	if (bExpandStrings && dwType == REG_EXPAND_SZ)
	{
		DWORD dwRequired = ExpandEnvironmentStrings(strValue, NULL, 0);
		if (dwRequired == 0)
			return GetLastError();

		CString strCopy = strValue;
		LPTSTR	lpszStr = strValue.GetBufferSetLength(dwRequired + 1);
		dwRequired = ExpandEnvironmentStrings(strCopy, lpszStr, dwRequired + 1);
		strValue.ReleaseBuffer();

		if (dwRequired == 0)
			return GetLastError();
	}
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::QueryValue(BYTE* pData, LPCTSTR lpszValueName, LPDWORD lpdwSize)
/////////////////////////////////////////////////////////////////////////////
{
	DWORD dwType = REG_BINARY;
	
	LONG lRes = ::RegQueryValueEx(m_hKey, (LPTSTR)lpszValueName, NULL, &dwType, pData, lpdwSize);
	
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG WINAPI CRegistryHelper::SetValue(HKEY hKeyParent, LPCTSTR lpszKeyName, LPCTSTR lpszValue, LPCTSTR lpszValueName)
/////////////////////////////////////////////////////////////////////////////
{
	ASSERT(lpszValue != NULL);
	CRegistryHelper key;
	LONG lRes = key.Create(hKeyParent, lpszKeyName);
	if (lRes == ERROR_SUCCESS)
		lRes = key.SetValue(lpszValue, lpszValueName);
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::SetKeyValue(LPCTSTR lpszKeyName, LPCTSTR lpszValue, LPCTSTR lpszValueName)
/////////////////////////////////////////////////////////////////////////////
{
	ASSERT(lpszValue != NULL);
	CRegistryHelper key;
	LONG lRes = key.Create(m_hKey, lpszKeyName);
	if (lRes == ERROR_SUCCESS)
		lRes = key.SetValue(lpszValue, lpszValueName);
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::RecurseDeleteKey(LPCTSTR lpszKey)
/////////////////////////////////////////////////////////////////////////////
// RecurseDeleteKey is necessary because on NT RegDeleteKey doesn't work 
// if the specified key has subkeys
{
	CRegistryHelper key;
	LONG lRes = key.Open(m_hKey, lpszKey);
	if (lRes != ERROR_SUCCESS)
		return lRes;
	FILETIME time;
	TCHAR szBuffer[256];
	DWORD dwSize = 256;
	while (RegEnumKeyEx(key.m_hKey, 0, szBuffer, &dwSize, NULL, NULL, NULL,
		&time)==ERROR_SUCCESS)
	{
		lRes = key.RecurseDeleteKey(szBuffer);
		if (lRes != ERROR_SUCCESS)
			return lRes;
		dwSize = 256;
	}
	key.Close();
	return DeleteSubKey(lpszKey);
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::RegEnumStringValues(CStringArray& strarrayValueNames, 
								  CStringArray* pstrarrayValues)
/////////////////////////////////////////////////////////////////////////////
{
	DWORD	cValues;
	DWORD	cbMaxValue;
	DWORD	cbMaxValueName;
	LONG	lRes = ::RegQueryInfoKey(m_hKey, NULL, NULL, 0, NULL, NULL, NULL,
						&cValues, &cbMaxValueName, &cbMaxValue, NULL, NULL);
	if (lRes != ERROR_SUCCESS)
		return lRes;

	TCHAR*	pszValueName	= NULL;
	TCHAR*	pszValue		= NULL;
	
	try
	{
		pszValueName	= new TCHAR[cbMaxValueName + 2];
		pszValue		= new TCHAR[cbMaxValue + 2];

		for (DWORD dwIndex = 0; dwIndex < cValues && lRes == ERROR_SUCCESS; dwIndex++)
		{
			DWORD	cbValueName = cbMaxValueName + 1;
			DWORD	cbValue		= cbMaxValue + 1;
			DWORD	dwType;

			if ((lRes = ::RegEnumValue(m_hKey, dwIndex, pszValueName, 
					&cbValueName, 0, &dwType, (LPBYTE)pszValue, &cbValue)) == ERROR_SUCCESS)
			{
				if (dwType == REG_SZ && cbValueName > 0)
				{
					strarrayValueNames.Add(pszValueName);
					if (pstrarrayValues)
						pstrarrayValues->Add(pszValue);
				}
			}
		}
	}
	catch(CMemoryException* e)
	{
		e->Delete();
		delete pszValueName;
		delete pszValue;
		AfxThrowMemoryException();
	}

	delete pszValueName;
	delete pszValue;
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
LONG CRegistryHelper::GetCountChildren(DWORD& dwSubKeys, DWORD& dwValues)
/////////////////////////////////////////////////////////////////////////////
{
	return ::RegQueryInfoKey(m_hKey, NULL, NULL, NULL, &dwSubKeys, NULL,
				NULL, &dwValues, NULL, NULL, NULL, NULL);
}

