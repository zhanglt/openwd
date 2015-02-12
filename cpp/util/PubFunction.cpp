#include "stdafx.h"
#include <iostream>
#include "util/InfoZip.h"
#include "afxdlgs.h"
#include "util/Regedit.h"
#include "util/BrowseDirDialog.h"
#include "util/ShowMsgDlg.h"
#include "util/PubFunction.h"

using namespace std;

#define DOWNLOAD 1   //����
#define SEND     0   //�ϴ�


#define NSIZE 11
//CString Dir[NSIZE]={"","���ı༭","���Ķ���","�����Ű�","���ĸ���","����༭","���涨��","�����Ű�","�������","��������","��������"};
CString Dir[NSIZE] = { "", "����༭", "���Ķ���", "�����Ű�", "���ĸ���", "����༭", "��������", "�����Ű�", "�������", "��������", "��������" };

CString szFileID = "SDopenwd";     //�ļ�ID��
CString szTmpID = "Tmp";
CString szFinalFile = "NOFILE";


CString lpTitle = "";  //�ڶ��߳���ʹ��
int CmdShow = 0;       //�ڶ��߳���ʹ��

CString szA_Name = "openwd.txt";


BOOL SetIpAndPort(char * Ip/*IP��ַ*/, int Port/*�˿�*/, char *ServerURL/*���������URL*/, char *Password/*��������*/)
{
	//MessageBox(NULL,Password,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	CString szValue;
	szValue = Ip;
	AfxGetApp()->WriteProfileString("openwd", "Ip", szValue);
	szValue.Format("%d", Port);
	AfxGetApp()->WriteProfileString("openwd", "Port", szValue);
	szValue = ServerURL;
	AfxGetApp()->WriteProfileString("openwd", "ServerURL", szValue);
	szValue = Password;
	AfxGetApp()->WriteProfileString("openwd", "Password", szValue);
	return true;
}

BOOL GetIpAndPort(char * Ip/*IP��ַ*/, int * Port/*�˿�*/, char *ServerURL/*���������URL*/){
	CString szValue;
	szValue = AfxGetApp()->GetProfileString("openwd", "Ip", "");
	strcpy(Ip, szValue);
	//AfxMessageBox(szValue);
	szValue = AfxGetApp()->GetProfileString("openwd", "Port", "");
	*Port = atoi(szValue);
	szValue = AfxGetApp()->GetProfileString("openwd", "ServerURL", "");
	strcpy(ServerURL, szValue);
	return true;
}

void SetID(char* FileID, char* TmpID){
	szFileID = FileID;
	szTmpID = TmpID;
	CreateDir();   //����Ŀ¼
}


BOOL GetTheCabarcFile()
{
	CString szPath[2];
	CString szInfo[2] = { "zipdll.dll", "cnzdll.dll" };
	szPath[0] = GetSysDirectory() + "\\ZIPDLL.DLL";
	szPath[1] = GetSysDirectory() + "\\UNZDLL.DLL";

	for (int i = 0; i < 2; i++)
	{

		int  bExist = IsTheFileExist(szPath[i]);
		bool bRight = 0;
		if (bExist)
		{   //ȷ���ļ������Ƿ���ȷ
			FILE * pfile = NULL;
			pfile = fopen(szPath[i], "rb");
			if (pfile == NULL)
			{
				MessageBox(NULL, "ȷ����ѹ�����ߵ��ļ�����ʱ������������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
				return false;
			}
			if (GetFileLen(pfile) > 90000) bRight = 1;
			fclose(pfile);
		}
		if (!bRight)
		{
			int	 Port = 0;
			char Ip[20];
			memset(Ip, 0, sizeof(Ip));
			char ServerURL[256];
			memset(ServerURL, 0, sizeof(ServerURL));
			try
			{
				if (!GetIpAndPort(Ip, &Port, ServerURL)) return false;   //��ȡ�˿ڡ�IP��ַ��������������
			}
			catch (CException * e)
			{
				e->ReportError();
				return false;
			}

			CInternetSession INetSession;
			CHttpConnection *pHttpServer = NULL;
			CHttpFile* pHttpFile = NULL;

			FILE * pfile = NULL;      //������������ص���Ϣ
			try
			{
				INetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_DATA_SEND_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_DATA_RECEIVE_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_CONTROL_SEND_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT, 30 * 60 * 1000);

				INTERNET_PORT nport = Port;
				if (nport > 0)
					pHttpServer = INetSession.GetHttpConnection(Ip, nport);
				else
					pHttpServer = INetSession.GetHttpConnection(Ip);   //Lotus����Ҫ�˿ں�

				pHttpFile = pHttpServer->OpenRequest(CHttpConnection::HTTP_VERB_POST, ServerURL, NULL, 1,
					NULL, NULL, INTERNET_FLAG_DONT_CACHE);

				pHttpFile->SendRequestEx(10);
				pHttpFile->Write(szInfo[i], 10);
				if (!(pHttpFile->EndRequest()))
				{
					MessageBox(NULL, "��������������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
					INetSession.Close();
					return false;
				}

				char buf[1001];
				memset(buf, 0, sizeof(buf));

				pfile = fopen(szPath[i], "wb+");
				if (pfile == NULL)
				{
					if (pHttpFile != NULL)	delete pHttpFile;
					if (pHttpServer != NULL)	delete pHttpServer;
					INetSession.Close();
					MessageBox(NULL, "���ؼӽ�ѹ������ʧ�ܣ�������������æ�����Ժ�����!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
					DeleteFile(szPath[i]);
					return false;
				}
				DWORD AllCount = 0;
				for (;;)
				{
					int len = pHttpFile->Read(buf, sizeof(buf)-1);
					AllCount += len;
					if (len == 0) break;
					fwrite((void*)buf, 1, len, pfile);
					memset(buf, 0, sizeof(buf));
				}
				if (pfile) fclose(pfile);

				if (AllCount < 100)
				{
					if (pHttpFile != NULL)	delete pHttpFile;
					if (pHttpServer != NULL)	delete pHttpServer;
					INetSession.Close();
					DeleteFile(szPath[i]);
					MessageBox(NULL, "���ؼӽ�ѹ������ʱ��������û�з�����Ϣ�����Ժ�����!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
					return false;
				}

				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();

			}
			catch (CInternetException *pInetEx)
			{
				char msg[400];
				memset(msg, 0, sizeof(msg));
				pInetEx->GetErrorMessage(msg, sizeof(msg)-1);
				CString szError;
				szError.Format("%s�����ԣ�", msg);
				MessageBox(NULL, szError, "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
				pInetEx->Delete();
				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				if (pfile) fclose(pfile);
				INetSession.Close();
				DeleteFile(szPath[i]);
				return false;
			}

		}
	}
	return true;

}

/*
��ȡsystemĿ¼
*/
CString GetSysDirectory()
{
	CString strPath;
	char buf[256];
	memset(buf, 0, sizeof(buf));
	GetSystemDirectory(buf, sizeof(buf));
	strPath = buf;
	return strPath;
}

/*

*/
CString GetIniName(int index)
{
	CString szIniFileName = GetSysDirectory() + "\\openwd\\" + Dir[index] + "\\" + szFileID + ".ini";
	return szIniFileName;
}

/*
void ShowMessage(DWORD nErrorCode)
{
LPVOID lpMsgBuf;
FormatMessage(
FORMAT_MESSAGE_ALLOCATE_BUFFER |
FORMAT_MESSAGE_FROM_SYSTEM |
FORMAT_MESSAGE_IGNORE_INSERTS,
NULL,
nErrorCode,
MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
(LPTSTR) &lpMsgBuf,
0,
NULL
);
MessageBox( NULL,(LPCTSTR)lpMsgBuf, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION );
// Free the buffer.
LocalFree( lpMsgBuf );

}
*/


/*
ɾ���ļ�
*/
BOOL DeleteDataFile(CString szPath)
{
	CFileFind fd;
	if (fd.FindFile(szPath + "\\*.*", NULL))
	for (;;)
	{
		BOOL bfind = fd.FindNextFile();
		CString szfileName = fd.GetFilePath();
		DeleteFile(szfileName);
		if (!bfind) break;
	}
	return true;
}

/*
�������ļ�
*/
BOOL ReNameFile(CString szSourceFile, CString szDesFile)
{
	CFile cf;
	try
	{
		cf.Rename(szSourceFile, szDesFile);
	}
	catch (...)
	{
		MessageBox(NULL, "�ļ�����ʧ��!", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		return false;
	}
	return true;
}




/************************************************************************/
/* ɾ����ʱ�ļ����ڵ���ʱ�ļ�                                           */
/************************************************************************/
void DeleteDirFile(int index)
{
	CString szDir = GetSysDirectory() + "\\openwd\\" + Dir[index] + "\\";
	LPSTR name = "*.*";
	LPSTR CurrentPath = (LPTSTR)(LPCTSTR)szDir;

	//ɾ��ָ��·���µ�ָ���ļ���֧��ͨ���
	//name:��ɾ�����ļ���CurrentPath:�ҵ����ļ�·��
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	char szHome[MAX_PATH];
	//charszFile[MAX_PATH];
	DWORD RightWrong;
	//HDCMyDiaDC;
	DWORD NameLength;
	//��ǰ�ĳ���·��
	RightWrong = GetCurrentDirectory(MAX_PATH, szHome);
	RightWrong = SetCurrentDirectory(CurrentPath);
	//�������ִ��·����Ȼ�󣬰ѵ�ǰ·���趨Ϊ��Ҫ���ҵ�·��
	hSearch = FindFirstFile(name, &FileData);
	if (hSearch != INVALID_HANDLE_VALUE)
	{
		NameLength = lstrlen(FileData.cFileName);

		DeleteFile(FileData.cFileName);
		while (FindNextFile(hSearch, &FileData)){
			//����һ���ļ����ҵ�һ��ɾ��һ��
			NameLength = lstrlen(FileData.cFileName);
			DeleteFile(FileData.cFileName);
		}
		FindClose(hSearch);
		//�رղ��Ҿ��
	}
	RightWrong = SetCurrentDirectory(szHome);
}




void  DeleteFilesEx(int index)
{
	int nMark = atoi(GetString("Mark", GetIniName(index)));
	int nInMark = atoi(GetString("Mark", GetIniName(index)));
	if (nMark + nInMark>0) return;

	CFileFind fd;
	CString szDir = GetSysDirectory() + "\\openwd\\" + Dir[index] + "\\";

	if (fd.FindFile(szDir + "\\*.*", 0))
	{
		for (;;)
		{
			CString strFileName;
			if (!fd.FindNextFile())
			{
				strFileName = fd.GetFileName();
				DeleteFile(szDir + strFileName);
				break;
			}
			strFileName = fd.GetFileName();
			DeleteFile(szDir + strFileName);
		}
	}
}


void CreateDir()
{
	CFileFind fd;
	CString szDir = GetSysDirectory();
	szDir += "\\openwd";
	//������Ŀ¼
	if (!fd.FindFile(szDir + "\\*.*"))
		CreateDirectory(szDir, NULL);
	//������Ŀ¼
	CString szPath;
	for (int i = 0; i<NSIZE; i++)
	{
		szPath = szDir + "\\" + Dir[i];
		if (!fd.FindFile(szPath + "\\*.*"))
			CreateDirectory(szPath, NULL);
	}
}





BOOL IsTheFileOpen(CString szFileName)
{

	LPSECURITY_ATTRIBUTES lpSecurityAttributes = NULL;
	HANDLE hdl = CreateFile(
		szFileName,
		GENERIC_READ | GENERIC_WRITE,
		FILE_SHARE_READ,			// ����ģʽ
		lpSecurityAttributes,		// ָ��ȫ���Ե�ָ��
		//OPEN_ALWAYS,				
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,		//�ļ�����
		0							// �����ļ����Ե��ļ�����Ŀ���
		);

	if (hdl == INVALID_HANDLE_VALUE)
	{
		int nErrorCode = GetLastError(); //��ȡ����ֵ

		CloseHandle(hdl);
		if (nErrorCode == 32 || nErrorCode == 33) return 1;
	}
	if (hdl) CloseHandle(hdl);

	return false;
}



BOOL IsTheFileExist(CString szFileName)
{
	FILE * pfile = NULL;
	pfile = fopen(szFileName, "r");
	if (pfile == NULL) return false;
	fclose(pfile);
	return true;
}

//

int CreateFileName(CString szFileName)
{


	if (IsTheFileExist(szFileName))
	{
		if (IsTheFileOpen(szFileName))
		{
			MessageBox(NULL, "�ļ��Ѿ��򿪣���رո��ļ������ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return 1;
		}
		else return 0;
	}
	//�����������ɣ����˳�

	FILE * pf = NULL;
	try
	{
		pf = fopen(szFileName, "wb+");
		if (pf) fclose(pf);
	}
	catch (CException * e)
	{
		e->Delete();
		if (pf) fclose(pf);
		return -1;
	}

	return 0;

}


CString GetFileName(CString Suffix/*��׺*/, CString szName, int index /*˳��*/)
{

	CString Path;
	int n = 0;

	Path.Format("%s\\openwd\\%s\\OA%s%s.%s", GetSysDirectory(), Dir[index], szName, szFileID, Suffix);

	int rec = CreateFileName(Path);
	if (rec != 0)  Path = "";

	return Path;
}

CString GetFile(CString Suffix/*��׺*/, CString szName, int index /*˳��*/)
{
	CString Path;
	int n = 0;

	Path.Format("OA%s%s.%s", szName, szFileID, Suffix);

	return Path;
}



/*

BOOL IsFileExist(int index)
{
if(IsTheFileExist(GetFileName("doc","D_",index))) return true;
if(IsTheFileExist(GetFileName("wps","T_",index))) return true;
return false;
}

*/

DWORD GetFileLen(FILE * pfile)
{
	DWORD nFileLen = 0;

	fseek(pfile, 0, SEEK_END);
	nFileLen = ftell(pfile);   //��ȡ�ļ�����
	rewind(pfile);           //ָ���Ƶ���ͷ
	return nFileLen;
}

DWORD GetFileLen(CString szFileName)
{
	FILE * pfile = NULL;
	DWORD nFileLen = 0;
	pfile = fopen(szFileName, "rb");
	if (pfile == NULL)
	{
		CString szInfo;
		szInfo.Format("�޷����ļ�%s��������!", szFileName);
		MessageBox(NULL, szInfo, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return -1;
	}
	fseek(pfile, 0, SEEK_END);
	nFileLen = ftell(pfile);   //��ȡ�ļ�����
	fclose(pfile);           //ָ���Ƶ���ͷ
	return nFileLen;
}




//pfile :�ļ�ָ��  2003/5/20 23:40
long FindString(FILE *pfile, char *szMark, BOOL bLast, DWORD nGap)
{

	fseek(pfile, 0, SEEK_END);   //��ȡ�ļ�����
	long len = ftell(pfile);
	rewind(pfile);

	fseek(pfile, nGap, SEEK_SET);

	int LenStart = strlen(szMark);  //��ȡ�����ַ�������
	char *lpbuf = new char[LenStart];
	memset(lpbuf, '\0', sizeof(lpbuf));
	long nPos = 0;
	long nNowPos = 0;
	while (nPos<len) //����
	{
		nPos = ftell(pfile);
		if (len - nPos<LenStart) {
			delete lpbuf;
			CString szstr;
			szstr.Format("û���ҵ��ַ���:%s", szMark);
			MessageBox(NULL, szstr, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return -1;
		}

		memset(lpbuf, 0, sizeof(lpbuf));
		fread(lpbuf, 1, LenStart, pfile);

		int i = 0;
		for (i; i<LenStart; i++)
		{
			if (lpbuf[i] != szMark[i])
				break;

		}


		if (i == LenStart) { nNowPos = nPos + LenStart; break; }

		fseek(pfile, 1 - LenStart, SEEK_CUR);


	}

	delete lpbuf;
	char ch = '\0';

	if (bLast)
	{

		fseek(pfile, -2 - LenStart, SEEK_CUR);
		long x = ftell(pfile);
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d || ch == 0x0a) nPos--;
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d) nPos--;

		return  nPos;
	}

	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d || ch == 0x0a) nNowPos++;
	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d) nNowPos++;

	rewind(pfile);

	return nNowPos;


}


BOOL  WriteToFile(CString szFileName, FILE * pfile, long BufferLen)
{
	FILE * pf = NULL;
	char buf[2];

	try
	{
		pf = fopen(szFileName, "wb+");
		if (pf == NULL)
		{
			MessageBox(NULL, "���������ļ�ʧ�ܣ������Ǹ��ļ����ڱ������ʹ�ã���������������ļ�!", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
			if (pf) fclose(pf);
			return false;
		}
		while (BufferLen>0)
		{
			memset(buf, 0, sizeof(buf));
			fread((void*)buf, 1, 1, pfile);
			fwrite((void*)buf, 1, 1, pf);
			BufferLen--;
		}
		if (pf) fclose(pf);

	}
	catch (CFileException *e)
	{
		char Msg[400];
		memset(Msg, 0, sizeof(Msg));
		e->GetErrorMessage(Msg, sizeof(Msg)-1);
		MessageBox(NULL, Msg, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		delete e;
		if (pf) fclose(pf);
		return false;
	}

	return true;
}


BOOL SplitFile(FILE *pfile, CString szFileName, char* szMarkStart, char* szMarkEnd)
{



	long nFirstLen;
	long nLastLen;
	long nFileLen;

	CString qq;


	//��Ѱ��LUKE��2004��4�£�
	qq = szMarkStart;
	if (qq == "PICTURESTART")
	{
		nFirstLen = FindStringAdd(pfile, szMarkStart);
		if (nFirstLen == -1) return false;
	}
	else
	{
		nFirstLen = FindString(pfile, szMarkStart);
		if (nFirstLen == -1) return false;
	}

	//��Ѱ��LUKE��2004��4�£�
	qq = szMarkEnd;
	if (qq == "DATAEND" || qq == "PICTUREEND")
	{
		nLastLen = FindStringAdd(pfile, szMarkEnd, true);
		if (nLastLen == -1) return false;
		nFileLen = nLastLen - nFirstLen;
	}
	else
	{
		nLastLen = FindString(pfile, szMarkEnd, true, nFirstLen);
		if (nLastLen == -1) return false;
		nFileLen = nLastLen - nFirstLen;
	}


	if (nFileLen>0)
	{
		fseek(pfile, nFirstLen, 0);
		if (!WriteToFile(szFileName, pfile, nFileLen)) { return false; }
	}  //С��0��=0˵��������

	return true;
}

bool OnFileCopy(CString m_SrcName, CString m_DstName)
{

	FILE * pfile = NULL;

	pfile = fopen(m_SrcName, "rb");
	if (pfile == NULL) return false;
	fseek(pfile, 0, SEEK_END);
	DWORD nlen = ftell(pfile);
	if (nlen == 0) return false;
	rewind(pfile);

	char * buf = new char[nlen];
	fread((void*)buf, 1, nlen, pfile);
	fclose(pfile);


	pfile = fopen(m_DstName, "wb+");
	if (pfile == NULL)
	{
		if (nlen>0) delete buf;
		buf = NULL;

		return false;
	}
	fwrite((buf), 1, nlen, pfile);
	if (pfile) fclose(pfile);
	if (nlen>0) delete buf;

	return true;
}



BOOL ExecuteFile(CString cmd, CString szFileName, UINT sw_cmd)
{
	LPVOID lpMsgBuf = NULL;

	HANDLE handle;
	STARTUPINFO si;
	PROCESS_INFORMATION pi;

	memset(&si, 0, sizeof(si));
	memset(&pi, 0, sizeof(pi));

	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = sw_cmd;  //Ĭ��ΪSW_SHOWNORMAL

	cmd += " " + szFileName;

	if (CreateProcess(NULL, cmd.GetBuffer(300), NULL, NULL, FALSE, CREATE_DEFAULT_ERROR_MODE, NULL,
		NULL, &si, &pi) == 0)
	{
		cmd.ReleaseBuffer();
		LPVOID lpMsgBuf;
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM |
			FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL,
			GetLastError(),
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR)&lpMsgBuf,
			0,
			NULL
			);
		// Display the string.
		if (sw_cmd>0)
		{
			cmd.ReleaseBuffer();
			CString szMsg;
			szMsg.Format("%s,��ѡ�������ķ�ʽ���ļ���", (LPCTSTR)lpMsgBuf);
			MessageBox(NULL, szMsg, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			// Free the buffer.
			LocalFree(lpMsgBuf);
		}
		return FALSE;
	}

	// �˴�������

	handle = OpenProcess(PROCESS_QUERY_INFORMATION | SYNCHRONIZE, FALSE, pi.dwProcessId);

	if (handle)
	{
		WaitForSingleObject(handle, INFINITE);
		CloseHandle(handle);
	}
	else if (sw_cmd>0)
	{
		MessageBox(NULL, "����ͬʱ������wPS�ļ���������ܻ������", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}


	return TRUE;

}



BOOL Compression(CString szCabFile, CString szSendFile)
{
	//	MessageBox( NULL,szSendFile, "������Ϣ", MB_OK | MB_ICONINFORMATION );

#define   FCOUNT 1

	char CabFile[256];
	memset(CabFile, 0, sizeof(CabFile));
	char SendFile[256];
//	char sendTxtFile[256];
	memset(SendFile, 0, sizeof(SendFile));
	strcpy(CabFile, szCabFile);
	strcpy(SendFile, szSendFile);
	char **pFiles = (char **) new  int[FCOUNT];
	//	for (int i=0; i<FCOUNT; i++)
	//	{
	//		pFiles[i] = new char[MAX_PATH+1];
	//		pFiles[i] = SendFile;
	//	}

	pFiles[0] = new char[MAX_PATH + 1];
	pFiles[0] = SendFile;

	//AfxMessageBox(pFiles[0]);

	//strcpy(sendTxtFile,szSendFile.Left(szSendFile.ReverseFind('.')+1)+"txt");

	//pFiles[1] = new char[MAX_PATH+1];

	//pFiles[1] ="c:/testc.txt" ; //sendTxtFile;





	DeleteFile(szCabFile);

	CInfoZip InfoZip;
	if (!InfoZip.InitializeZip())
	{
		MessageBox(NULL, "��ʼ��ѹ������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!InfoZip.AddFiles(CabFile, pFiles, FCOUNT))
	{
		return false;
	}

	if (!InfoZip.Finalize())
	{
		MessageBox(NULL, "����ѹ������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	return true;
}


////���ܣ���LWZ�㷨ѹ���ļ�,����CABARC ѹ��
//BOOL DeCompression(CString szCabFile,CString szFileName,int index)
//{   
//    FILE * pfile=NULL;
//	CString szPath;  // ������ʱ�ļ�
//  	CString szComDataFile;
//	CString szCommand;
//	pfile=fopen(szCabFile,"rb");
//	if(pfile == false) 
//	{
//	    MessageBox(NULL,"��ȡ�����ļ���ʧ�ܣ�������!","ϵͳ��Ϣ",MB_OK|MB_ICONERROR);
//		return false;
//	}
//	int nFileLen=GetFileLen(pfile);
//	if(nFileLen<2) 
//	{
//		if(pfile) fclose(pfile);
//		DeleteFile(szCabFile);  //ɾ��ѹ���ļ� 
//		return true;
//	}  //û�����ݵ��ļ�
//	
//	fseek(pfile,60,0);
//
//	char ch[2];
//	CString szName;
//	for(;;)
//	{
//		memset(ch,0,sizeof(ch));
//		fread(ch,1,1,pfile);
//		if(ch[0]==0) break;
//		szName+=ch;
//	}
//    if(pfile) fclose(pfile);
//
//    szComDataFile.Format("%s\\Zhonglu\\%s\\%s",GetSysDirectory(),Dir[index],szName);
//    szCommand.Format("%s\\Zhonglu\\%s",GetSysDirectory(),"cabarc.exe -o x");
//
//	szPath.Format(" %s\\Zhonglu\\%s\\",GetSysDirectory(),Dir[index]);
//	szPath =szCabFile+szPath;
//	if(!ExecuteFile(szCommand,szPath,SW_HIDE)) 
//	{
//		MessageBox(NULL,"ִ�н�ѹ����ʧ�ܣ�������!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
//		return false;
//	}//���ļ���ѹ�� 
//
//    DeleteFile(szCabFile);  //ɾ��ѹ���ļ� 
//	DeleteFile(szFileName);  //��ԭ���ļ�ɾ��
//	//����ѹ����ļ�����Ϊ�����ļ�
//	if(!ReNameFile(szComDataFile,szFileName)) return false;
//
//    return true;
//}


BOOL DeCompression(CString& szCabFile, CString szFileName, int index)
{
	FILE * pfile = NULL;
	CString szPath;
	CString szComDataFile;
	CString szCommand;
	pfile = fopen(szCabFile, "rb");
	// MessageBox(NULL,szCabFile,"szCabFile",MB_OK|MB_ICONERROR);
	if (pfile == false)
	{
		MessageBox(NULL, "��ȡ�����ļ���ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		return false;
	}
	int nFileLen = GetFileLen(pfile);
	if (nFileLen<2)
	{
		if (pfile) fclose(pfile);
		DeleteFile(szCabFile);
		return true;
	}

	fseek(pfile, 26, 0);
	char ch[2];
	memset(ch, 0, sizeof(ch));
	fread(ch, 1, 1, pfile);
	int count = ch[0];
	fseek(pfile, 30, 0);
	CString szName;
	for (;;)
	{
		memset(ch, 0, sizeof(ch));
		fread(ch, 1, 1, pfile);
		if (count<1) break;
		szName += ch;
		count--;
	}
	if (pfile) fclose(pfile);

	int nlen = szFileName.Find("##");

	if (index == 10 && nlen<0)
	{
		szPath = szFileName;
	}
	else
	{

		szComDataFile.Format("%s\\openwd\\%s\\%s", GetSysDirectory(), Dir[index], szName);

		szPath.Format("%s\\openwd\\%s\\", GetSysDirectory(), Dir[index]);
	}
	CInfoZip InfoZip;
	if (!InfoZip.InitializeUnzip())
	{
		MessageBox(NULL, "��ʼ����ѹ������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!InfoZip.ExtractFiles(szCabFile/*Ҫ��ѹ���ļ�*/, szPath/*Ŀ��·��*/))
	{
		return false;
	}

	if (!InfoZip.FinalizeUnzip())
	{
		MessageBox(NULL, "������ѹ������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}


	if (index == 10)
	{
		if (nlen>0)
		{
			szFileName.Replace("##", "");
			MessageBox(NULL, szComDataFile, "szComDataFile222222", MB_OK | MB_ICONINFORMATION);
			if (!OnFileCopy(szComDataFile, szFileName)) return false;
		}
		else
		{

			CString szAttachFileName = szPath + "\\" + szName;
			CString szDisName = szPath + "\\" + szA_Name;
			DeleteFile(szDisName);   //ɾ���Ѵ��ڵ��ļ�

			if (!ReNameFile(szAttachFileName, szDisName)) return false;
		}


	}
	else
	{
		if (szFileName.Find("openwdOA")<0)    DeleteFile(szCabFile);  //ɾ��ѹ���ļ� 
		DeleteFile(szFileName);  //��ԭ���ļ�ɾ��
		//����ѹ����ļ�����Ϊ�����ļ�
		if (!ReNameFile(szComDataFile, szFileName)) return false;

		szCabFile = szName;
	}
	return true;
}

CString GetFileEx(CString szPath)
{   //������ļ�������Ϊ��
	CString Ex[] = { "DOC", "XLS", "BMP", "PPT", "JPG", "GIF", "TXT", "INI", "LOG", "TIF", "MDB" };

	char path[256];
	memset(path, 0, sizeof(path));
	strcpy(path, szPath);
	CString szEx;
	for (int i = strlen(path) - 1; i>0; i--)
	{
		if (path[i] == '.') break;
		szEx = path[i] + szEx;
	}
	szEx.MakeUpper();
	for (int n = 0; n<7; n++)
	{
		if (szEx == Ex[n]) return Ex[n];
	}

	return "����";

}



BOOL OpenAttachment(CString szFileName)
{
	char buf[256];
	memset(buf, 0, sizeof(buf));
	GetSystemDirectory(buf, sizeof(buf));

	CString szExcuteFile;

	HKEY KEY = HKEY_LOCAL_MACHINE;
#define NCOUNT 2
	CString szKeyPath[NCOUNT] = {
		"SOFTWARE\\Microsoft\\Office\\9.0\\Word\\InstallRoot\\",
		"SOFTWARE\\Microsoft\\Office\\10.0\\Word\\InstallRoot\\"
	};
	CString szKeyValue = "Path";
	int count = 0;
	int i = 0;
	for (i; i<NCOUNT; i++)
	{
		if (!GetProfileString(KEY, szKeyPath[i], szKeyValue)) continue;
		else break;
	}

	if (i >= NCOUNT)
	{
		MessageBox(NULL, "�޷���ȡOffice 2000�İ�װ·������ȷ�Ϻ����ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	CString szEx;
	szEx = GetFileEx(szFileName);
	szEx.MakeUpper();  //�˴�Ƿ����
	if (szEx == "DOC")
	{
		szExcuteFile += szKeyValue + "wps.exe";
		if (!ExecuteFile(szExcuteFile, szFileName))	return false;
	}
	else if (szEx == "XLS")
	{
		szExcuteFile = szKeyValue + "EXCEL.EXE";
		if (!ExecuteFile(szExcuteFile, szFileName))	return false;
	}
	else if (szEx == "PPT")
	{
		szExcuteFile = szKeyValue + "POWERPNT.EXE";
		if (!ExecuteFile(szExcuteFile, szFileName))	return false;
	}
	else if (szEx == "MDB")
	{
		szExcuteFile = szKeyValue + "MSACCESS.EXE";
		if (!ExecuteFile(szExcuteFile, szFileName))	return false;

	}
	else if (szEx == "BMP")
	{
		szExcuteFile = "MSPAINT.exe";
		if (!ExecuteFile(szExcuteFile, szFileName))	return false;
	}
	else if (szEx == "TXT" || szEx == "INI" || szEx == "LOG")
	{
		szExcuteFile = "NotePad.exe";
		if (!ExecuteFile(szExcuteFile, szFileName)) return false;
	}
	else
	{
		if (IDYES != MessageBox(NULL, "ϵͳ��֧�ֱ༭�����ļ������Ƿ�Ҫ�����", "ϵͳ��Ϣ", MB_YESNO | MB_ICONINFORMATION))
			return false;

		int rec = (int)ShellExecute(NULL, "open", szFileName, NULL, NULL, SW_SHOWNORMAL);
		if (rec <= 32)
		{
			MessageBox(NULL, "û�к��ʵ�Ӧ�������򿪴����ļ����밲װ��Ӧ��Ӧ�ó�������ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		}
		return false;

	}
	return true;
}











void SetFlagEx(int index, int nflag)
{
	CString szMark;
	szMark.Format("%d", nflag);
	WriteString("Mark", szMark, GetIniName(index));
}
int GetFlagEx(int index)
{
	CString szMark = GetString("Mark", GetIniName(index));
	if (szMark == "") return -1;
	else return atoi(szMark);
}

void SetInSureFlagEx(int index, int nflag)
{
	CString szMark;
	szMark.Format("%d", nflag);
	WriteString("InMark", szMark, GetIniName(index));
}

int GetInSureFlagEx(int index)
{
	CString szMark = GetString("InMark", GetIniName(index));
	if (szMark == "") return -1;
	else return atoi(szMark);
}

UINT ShowWindowEx(LPVOID pParam)
{
	//2003/8/24  10:04  ��ĳ�˻����Ĵ����޷����ڶ��㣬����Ҫд��־ȷ��ԭ��
	HWND  hWnd = 0;
	CString szTitle = lpTitle;
	if (CmdShow<0)
	{
		CString szTemp;
		szTitle = szFinalFile + " - Kingsoft wps";
		WriteLog(szTitle);
		for (int i = 0; i<100; i++)
		{
			hWnd = ::FindWindow(NULL, szTitle);

			szTemp.Format("ѭ����hWnd=%d", hWnd);
			WriteLog(szTemp);
			if (hWnd) break;
			Sleep(100);
		}
		int rec = ::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		szTemp.Format("���ô���ʱ�ķ���ֵ=%d", rec);
		WriteLog(szTemp);
	}
	else  // ������ҳ
	{
		hWnd = ::FindWindow(NULL, szTitle);
		::ShowWindow(hWnd, CmdShow);
	}

	return 1;
}

void ShowWinEx(CString szTitle, int nCmdShow)
{  //2003/8/24  10:04  ��ĳ�˻����Ĵ����޷����ڶ��㣬����Ҫд��־ȷ��ԭ��
	//2004/1/7��ĳЩ������wps���ٶȽ��������޷��������ڶ��㣬�ʽ���δ����Ϊ���߳�
	//ѭ�����ҷ�ʽ
	HWND  hWnd = 0;
	if (nCmdShow<0)   //�����������wps
	{
		CString szTemp;
		szTitle = szFinalFile + " - Kingsoft wps";
		//	WriteLog(szTitle);
		for (int i = 0; i<10; i++)
		{
			hWnd = ::FindWindow(NULL, szTitle);
			if (hWnd) break;  //�ҵ����˳�
			szTemp = szTitle;
			//AfxMessageBox(szTemp);
			//szTemp.Replace(".doc","");
			//AfxMessageBox(szTemp);
			//	WriteLog(szTemp);
			hWnd = ::FindWindow(NULL, szTemp);
			if (hWnd) break;  //�ҵ����˳�
			szTemp.Format("ѭ����hWnd=%d", hWnd);

			//  WriteLog(szTemp);

			Sleep(200);
		}
		int rec = ::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
		//	szTemp.Format("���ô���ʱ�ķ���ֵ=%d",rec);
		//	WriteLog(szTemp);

	}
	else  // ������ҳ
	{
		//	WriteLog("������ҳ");
		hWnd = ::FindWindow(NULL, szTitle);
		::ShowWindow(hWnd, nCmdShow);
	}
	//2004/1/9  
	//    szTitle=lpTitle;
	//    CmdShow = nCmdShow;
	// 	AfxBeginThread(ShowWindowEx,NULL);

}

//�ֽ��ļ���
void GetAllFileNames(CStringArray &szItems, CString szFileNames)
{
	CString szTemp;
	int nlen = -1;
	nlen = szFileNames.Find("#|#");
	while (nlen>-1)
	{
		szTemp = szFileNames.Left(nlen);
		if (szTemp != "")	szItems.Add(szTemp);
		szFileNames = szFileNames.Mid(nlen + 3);
		nlen = szFileNames.Find("#|#");
		if (nlen<0) break;

	}
	if (szFileNames != "")	szItems.Add(szFileNames);

}

BOOL JudgeFileIgnoreOrNot(CString szPath, CString & szFileName)
{

	if (IsTheFileExist(szPath + "\\" + szFileName))
	{
		ShowMsgDlg dlg;
		dlg.szFileName = szFileName;
		dlg.m_Path = szPath;
		dlg.DoModal();
		if (dlg.nMark == 1)
		{
			szFileName = szPath;
		}
		else if (dlg.nMark == 2)
		{
			szFileName = szPath + "##" + "\\" + dlg.szFileName;
		}
		else
		{
			return false;
		}
	}
	else
	{
		szFileName = szPath;
	}
	return true;
}


//��Ѱ������LUKE��2004��4�£�
long FindStringAdd(FILE *pfile, char *szMark, BOOL bLast, DWORD nGap)
{
	fseek(pfile, 0, SEEK_END);   //��ȡ�ļ�����
	long len = ftell(pfile);

	int LenStart = strlen(szMark);  //��ȡ�����ַ�������
	char *lpbuf = new char[LenStart];
	memset(lpbuf, '\0', sizeof(lpbuf));

	long nNowPos = len - 1;
	fseek(pfile, 0 - LenStart, SEEK_END);
	long nPos = ftell(pfile);

	while (nPos >= 0) //����
	{
		nPos = ftell(pfile);

		memset(lpbuf, 0, sizeof(lpbuf));
		fread(lpbuf, 1, LenStart, pfile);

		int i = 0;
		for (i; i<LenStart; i++)
		{
			if (lpbuf[i] != szMark[i])
				break;

		}


		if (i == LenStart) { nNowPos = nPos + LenStart; break; }

		if (nPos = 0)
		{
			delete lpbuf;
			CString szstr;
			szstr.Format("û���ҵ��ַ���:%s", szMark);
			MessageBox(NULL, szstr, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return -1;
		}

		fseek(pfile, -1 - LenStart, SEEK_CUR);


	}

	delete lpbuf;
	char ch = '\0';

	if (bLast)
	{
		//ȥ���س����з�
		fseek(pfile, -2 - LenStart, SEEK_CUR);
		long x = ftell(pfile);
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d || ch == 0x0a) nPos--; //�������س�
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d) nPos--; //���������з�

		return  nPos;
	}

	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d || ch == 0x0a) nNowPos++; //�������س�
	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d) nNowPos++; //���������з�

	rewind(pfile);   //���ļ�ָ���λ

	return nNowPos;  //�����һλ��


}