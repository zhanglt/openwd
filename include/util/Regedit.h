/************************************************************************/
/*×¢²á±í²Ù×÷
*************************************************************************/

BOOL GetProfileString(HKEY KEY, CString szPath, CString &szKeyValue);

BOOL CreateKey(HKEY KEY, CString szPath, CString szValue);

BOOL WriteProfileString(HKEY KEY, CString szPath, BYTE *pcszValue);

CString GetString(CString szKeyValue, CString szFileName);

void WriteString(CString szKeyValue, CString szValue, CString szFileName);

void WriteLog(CString szText);

void SetKey(HKEY KEY);

BOOL Open(LPCTSTR lpSubKey);

BOOL CreateKey(LPCTSTR lpSubKey);

void Close();

BOOL Write(LPCTSTR lpValueName, LPCTSTR lpValue);
