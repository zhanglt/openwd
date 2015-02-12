/*
注册表操作
*/

#include "stdafx.h"
#include  <fstream>
#include "util/Regedit.h"

using namespace std;

HKEY m_hKey = HKEY_CLASSES_ROOT;


BOOL GetProfileString(HKEY KEY, CString szPath, CString &szKeyValue)
{
	HKEY hKEY; 

	long ret = ::RegOpenKeyEx(KEY, szPath, 0, KEY_READ, &hKEY);

	if (ret != ERROR_SUCCESS) return false;

	LPBYTE owner_Get = new BYTE[256]; 

	DWORD type_1 = REG_SZ; 

	DWORD cbData_1 = 80;

	ret = ::RegQueryValueEx(hKEY, szKeyValue, NULL, &type_1, owner_Get, &cbData_1);

	if (ret != ERROR_SUCCESS) {
		::RegCloseKey(hKEY); szKeyValue = ""; 
	}  

	else szKeyValue = CString(owner_Get);

	delete[] owner_Get;

	::RegCloseKey(hKEY); 

	return true;
}


//设置主键
void SetKey(HKEY KEY)
{
	m_hKey = KEY;
}

BOOL Open(LPCTSTR lpSubKey)
{
	ASSERT(m_hKey);
	ASSERT(lpSubKey);

	HKEY hKey;
	long lReturn = RegOpenKeyEx(m_hKey, lpSubKey, 0L, KEY_ALL_ACCESS, &hKey);

	if (lReturn == ERROR_SUCCESS)
	{
		m_hKey = hKey;
		return TRUE;
	}
	return FALSE;

}

BOOL CreateKey(LPCTSTR lpSubKey)
{
	ASSERT(m_hKey);
	ASSERT(lpSubKey);

	HKEY hKey;
	DWORD dw;
	long lReturn = RegCreateKeyEx(m_hKey, lpSubKey, 0L, NULL, REG_OPTION_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &dw);

	if (lReturn == ERROR_SUCCESS)
	{
		m_hKey = hKey;
		return TRUE;
	}

	return FALSE;

}

BOOL Write(LPCTSTR lpValueName, LPCTSTR lpValue)
{
	long lReturn = RegSetValueEx(m_hKey, lpValueName, 0L, REG_SZ, (const BYTE *)lpValue, strlen(lpValue) + 1);

	if (lReturn == ERROR_SUCCESS)
		return TRUE;

	return FALSE;

}

void Close()
{
	if (m_hKey)
	{
		RegCloseKey(m_hKey);
		m_hKey = NULL;
	}

}


//获取配置项的值
CString GetString(CString szKeyValue, CString szFileName)
{
	char buf[256];

	memset(buf, 0, sizeof(buf));

	GetPrivateProfileString(
		"openWDcom",
		szKeyValue,
		"",
		buf,
		sizeof(buf)-1,
		szFileName
		);
	return buf;
}

void WriteString(CString szKeyValue, CString szValue, CString szFileName)
{
	WritePrivateProfileString("Telecom", szKeyValue, szValue, szFileName);
}


void WriteLog(CString szText)
{
	CString szFileName;

	char buf[256];

	memset(buf, 0, sizeof(buf));

	GetSystemDirectory(buf, sizeof(buf));

	szFileName = buf;

	szFileName += "\\openwd\\log.txt";

	ofstream wf(szFileName , ios_base::out | ios_base::app | ios_base::binary);

	if (wf.is_open()){
		wf << szText << endl;
		wf.close();
	}else{
		
		MessageBox(NULL, "Write Log Error!", "SYSTEM", MB_OK | MB_ICONINFORMATION);
		return;
	}




}
