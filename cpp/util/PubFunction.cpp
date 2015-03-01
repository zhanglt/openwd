#include "stdafx.h"
#include <iostream>

#include "util/InfoZip.h"
#include "afxdlgs.h"
#include "util/Regedit.h"
#include "util/BrowseDirDialog.h"
#include "util/ShowMsgDlg.h"
#include "util/PubFunction.h"

using namespace std;

#define DOWNLOAD 1   //下载
#define SEND     0   //上传
#define NSIZE 11
//CString Dir[NSIZE]={"","正文编辑","正文定稿","正文排版","正文盖章","传真编辑","传真定稿","传真排版","传真盖章","附件处理","附件下载"};
CString Dir[NSIZE] = { "", "文书编辑", "正文定稿", "正文排版", "正文盖章", "传真编辑", "生成文书", "传真排版", "文书盖章", "附件处理", "附件下载" };

CString szFileID = "SDopenwd";     //文件ID号
CString szTmpID = "Tmp";
CString szFinalFile = "NOFILE";


CString lpTitle = "";  //在多线程中使用
int CmdShow = 0;       //在多线程中使用
CString szA_Name = "openwd.txt";
//************************************
// Method:    DocConnectionHttp
// FullName:  DocConnectionHttp
// Access:    public 
// Returns:   BOOL
// Qualifier:
// Parameter: char * TextBuf
// Parameter: DWORD nFileLen
// Parameter: int index
// Parameter: int bDownLoad
// Parameter: CString szAttachmentFileName
//************************************
BOOL DocConnectionHttp(CString TextBuf, DWORD nFileLen, int index, int bDownLoad, CString szAttachmentFileName)
{
	
	if (bDownLoad)   //>0 表示下载
	{
		if (!GetTheCabarcFile()) return false; //下载加解压缩工具
		/*
		int rec = IsNeedLoad(index);
		if (rec == -1) return false;             //出错
		if (rec == 0) return true;               //已经下载
		*/
	}
	CString Ip, Port, ServerURL;
	try
	{

		if (!GetIpAndPort(Ip, Port, ServerURL)) {  //获取端口、IP地址、及服务器名称
			return false;
		}
		/*
		if (AfxGetApp()->GetProfileString("Telecom", "Large", "") == "1")
		{
		memset(ServerURL, 0, sizeof(ServerURL));
		strcpy(ServerURL, "servlet/ULoadBDoc");
		}*/
	}
	catch (CException * e)
	{
		e->ReportError();
		return false;
	}

	CInternetSession INetSession;
	CHttpConnection *pHttpServer = NULL;
	CHttpFile       *pHttpFile = NULL;

	FILE * pfile = NULL;      //保存服务器下载的信息
	CString szPath;         //保存临时文件
	szPath.Format("%s\\openwd\\%s\\TempDoc.dat", GetSysDirectory(), Dir[index]);
	try
	{
		INetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT,			30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_SEND_TIMEOUT,		30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_RECEIVE_TIMEOUT,		30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_SEND_TIMEOUT,		30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT,	30 * 60 * 1000);
		
		INTERNET_PORT nport = atoi(Port);
		if (nport>0)
			pHttpServer = INetSession.GetHttpConnection(Ip, nport);
		else
			pHttpServer = INetSession.GetHttpConnection(Ip);

		pHttpFile = pHttpServer->OpenRequest(CHttpConnection::HTTP_VERB_POST, ServerURL, NULL, 1, NULL, NULL, INTERNET_FLAG_DONT_CACHE);
		pHttpFile->SendRequestEx(nFileLen);
		pHttpFile->Write(TextBuf, nFileLen);


		
		if (!(pHttpFile->EndRequest()))
		{
			MessageBox(NULL, "服务器结束请求失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
			INetSession.Close();
			return false;
		}
		
		char buf[1001];
		memset(buf, 0, sizeof(buf));
		if (bDownLoad)
		{
			pfile = fopen(szPath, "wb");
			if (pfile == NULL)
			{
				if (pHttpFile != NULL)		delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "无法生成临时下载文件，可能是网络正忙，请稍后重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
				return false;
			}
			
			DWORD AllCount = 0;
			for (;;)
			{
				int len = pHttpFile->Read(buf, sizeof(buf)-1);
				AllCount += len;
				if (len == 0) break;							  //将服务器返回信息息全部读出
				fwrite((void*)buf, 1, len, pfile);
				memset(buf, 0, sizeof(buf));
			}   //保存文件结束 
			if (pfile) fclose(pfile);
			CString szStr;
			szStr = buf;

			if (szStr == "large"){
				if (pHttpFile != NULL)		delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "文件太大，无法进行编辑操作!", "系统信息", MB_OK | MB_ICONINFORMATION);
				return false;
			}
			
			if (AllCount<100){
				if (pHttpFile != NULL)		delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "服务器没有返回信息，请稍后重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
				return false;

			}

		}
		else
		{
			CString sztemp;

			bool issuccessed = false;
			int findposition = 0;

			int len = pHttpFile->Read(buf, sizeof(buf)-1);    //从端口读取返回信息
			sztemp = buf;
			sztemp.MakeUpper();

			while (findposition<len) //查找
			{
				if (len - findposition<2)
				{
					issuccessed = false;
					break;
				}

				int i = 0;
				for (i; i<2; i++)
				{
					int j;
					char tempmark[4] = "OK";
					j = findposition + i;

					if (sztemp[j] != tempmark[i])
						break;

				}

				if (i == 2) { issuccessed = true; break; }

				findposition = findposition + 1;
			}

			if (issuccessed == false)
			{
				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "上传文件失败，请重新提交!", "系统信息", MB_OK | MB_ICONINFORMATION);
				return false;
			}
		}
		//释放内存空间
		if (pHttpFile != NULL)		delete pHttpFile;
		if (pHttpServer != NULL)	delete pHttpServer;
		INetSession.Close();		
	}
	catch (CInternetException *pInetEx)
	{   //释放内存空间
		char msg[400];
		memset(msg, 0, sizeof(msg));
		pInetEx->GetErrorMessage(msg, sizeof(msg)-1);
		CString szError;
		szError.Format("%s请重试！", msg);
		MessageBox(NULL, szError, "系统信息", MB_OK | MB_ICONERROR);
		pInetEx->Delete();
		if (pHttpFile != NULL)	delete pHttpFile;
		if (pHttpServer != NULL)	delete pHttpServer;
		if (pfile) fclose(pfile);
		INetSession.Close();
		return false;
	}

	if (bDownLoad)
	{
		if (!MakeFile(szPath, index, szAttachmentFileName)) return false;

	}

	return true;
}
int  IsNeedLoad(int index)
{
	int nMark = atoi(GetString("Mark", GetIniName(index)));
	int nInMark = atoi(GetString("Mark", GetIniName(index)));
	if (nMark + nInMark <= 0) DeleteDirFile(index);    //DeleteAll(index);

	//判断下列文件是否打开,当文件名为空时，说明文件已经打开
	if (GetFileName("ini", "P_", index) == "") return -1;  //权限
	if (GetFileName("doc", "H_", index) == "") return -1;  //头文件
	if (GetFileName("doc", "T_", index) == "") return -1;  //模板
	if (GetFileName("doc", "D_", index) == "") return -1;  //数据文件
	if (GetFileName("bmp", "B_", index) == "") return -1;  //公章

	//如果不清理文件，则每次都要下载
	if (AfxGetApp()->GetProfileString("Telecom", "DeleteAllFile", "") != "") return true;

	if (GetString("IsNeedLoad", GetIniName(index)) == "1") return false;  //如果存在则不需要下载

	return true;
}

BOOL MakeFile(CString szFileName, int index, CString szAttachmentPath)
{

	FILE *pfile = NULL;
	pfile = fopen(szFileName, "rb");
	if (pfile == NULL)
	{
		MessageBox(NULL, "打开已下载的数据文件失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!SplitFile(pfile, GetFileName("ini", "P_", index), "HEADSTART", "HEADEND")) { fclose(pfile); return false; }

	if (!SplitFile(pfile, GetFileName("zip", "H_", index), "FILEHEADSTART", "FILEHEADEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "T_", index), "TMPSTART", "TMPEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "D_", index), "DATASTART", "DATAEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "B_", index), "PICTURESTART", "PICTUREEND")) { fclose(pfile); return false; }
	if (pfile) fclose(pfile);

	if (index == 10)  //2003/11/26  添加了下载所有附件的功能
	{  //下载所有附件
		if (!DeCompression(GetFileName("zip", "D_", index), szAttachmentPath, index)) return false;
	}
	else
	{

		//解压缩文件
		if (!DeCompression(GetFileName("zip", "H_", index), GetFileName("doc", "H_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "T_", index), GetFileName("doc", "T_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "D_", index), GetFileName("doc", "D_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "B_", index), GetFileName("bmp", "B_", index), index)) return false;

	}
	return true;
}
BOOL SetIpAndPort(CString Ip/*IP地址*/, CString Port/*端口*/, CString ServerURL/*请求服务器URL*/, CString Password/*解锁密码*/)
{
	//MessageBox(NULL,Password,"系统信息",MB_OK|MB_ICONINFORMATION);
	//CString szValue;
	//szValue = Ip;
	AfxGetApp()->WriteProfileString("openwd", "Ip", Ip);
	//szValue.Format("%d", Port);
	AfxGetApp()->WriteProfileString("openwd", "Port", Port);
	//szValue = ServerURL;
	AfxGetApp()->WriteProfileString("openwd", "ServerURL", ServerURL);
	//szValue = Password;
	AfxGetApp()->WriteProfileString("openwd", "Password", Password);
	return true;
}

BOOL GetIpAndPort(CString &Ip/*IP地址*/, CString &Port/*端口*/, CString &ServerURL/*请求服务器URL*/){
	

	::GetProfileString("openwd", "ServerURL", "jc/legalDoc", ServerURL.GetBuffer(50), 50);


	::GetProfileString("openwd", "Port", "80", Port.GetBuffer(6), 6);


	::GetProfileString("openwd", "Ip", "127.0.0.1", Ip.GetBuffer(15), 15);
	if (&Ip == NULL || &Port == NULL || &ServerURL == NULL)
	{
		return false;
	}

	return true;
}

void SetID(char* FileID, char* TmpID){
	szFileID = FileID;
	szTmpID = TmpID;
	CreateDir();   //创建目录
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
		{   //确定文件长度是否正确
			FILE * pfile = NULL;
			pfile = fopen(szPath[i], "rb");
			if (pfile == NULL)
			{
				MessageBox(NULL, "确定解压缩工具的文件长度时出错，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
				return false;
			}
			if (GetFileLen(pfile) > 90000) bRight = 1;
			fclose(pfile);
		}
		if (!bRight)
		{/*
			int	 Port = 0;
			char Ip[20];
			memset(Ip, 0, sizeof(Ip));
			char ServerURL[256];
			memset(ServerURL, 0, sizeof(ServerURL));*/
			CString Ip, Port, ServerURL;
			try
			{
				if (!GetIpAndPort(Ip, Port, ServerURL)) return false;   //获取端口、IP地址、及服务器名称
			}
			catch (CException * e)
			{
				e->ReportError();
				return false;
			}

			CInternetSession INetSession;
			CHttpConnection *pHttpServer = NULL;
			CHttpFile* pHttpFile = NULL;

			FILE * pfile = NULL;      //保存服务器下载的信息
			try
			{
				INetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_DATA_SEND_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_DATA_RECEIVE_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_CONTROL_SEND_TIMEOUT, 30 * 60 * 1000);
				INetSession.SetOption(INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT, 30 * 60 * 1000);

				INTERNET_PORT nport = atoi(Port);
				if (nport > 0)
					pHttpServer = INetSession.GetHttpConnection(Ip, nport);
				else
					pHttpServer = INetSession.GetHttpConnection(Ip);   //Lotus不需要端口号

				pHttpFile = pHttpServer->OpenRequest(CHttpConnection::HTTP_VERB_POST, ServerURL, NULL, 1,
					NULL, NULL, INTERNET_FLAG_DONT_CACHE);

				pHttpFile->SendRequestEx(10);
				pHttpFile->Write(szInfo[i], 10);
				if (!(pHttpFile->EndRequest()))
				{
					MessageBox(NULL, "服务器结束请求失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
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
					MessageBox(NULL, "下载加解压缩工具失败，可能是网络正忙，请稍后重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
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
					MessageBox(NULL, "下载加解压缩工具时，服务器没有返回信息，请稍后重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
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
				szError.Format("%s请重试！", msg);
				MessageBox(NULL, szError, "系统信息", MB_OK | MB_ICONERROR);
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
获取system目录
*/
CString GetSysDirectory()
{
	CString strPath;
	char buf[256];
	memset(buf, 0, sizeof(buf));
	GetWindowsDirectory(buf, sizeof(buf));
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
MessageBox( NULL,(LPCTSTR)lpMsgBuf, "系统信息", MB_OK | MB_ICONINFORMATION );
// Free the buffer.
LocalFree( lpMsgBuf );

}
*/


/*
删除文件
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
重命名文件
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
		MessageBox(NULL, "文件更名失败!", "系统信息", MB_OK | MB_ICONERROR);
		return false;
	}
	return true;
}




/************************************************************************/
/* 删除临时文件夹内的临时文件                                           */
/************************************************************************/
void DeleteDirFile(int index)
{
	CString szDir = GetSysDirectory() + "\\openwd\\" + Dir[index] + "\\";
	LPSTR name = "*.*";
	LPSTR CurrentPath = (LPTSTR)(LPCTSTR)szDir;

	//删除指定路径下的指定文件，支持通配符
	//name:被删除的文件；CurrentPath:找到的文件路径
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	char szHome[MAX_PATH];
	//charszFile[MAX_PATH];
	DWORD RightWrong;
	//HDCMyDiaDC;
	DWORD NameLength;
	//当前的程序路径
	RightWrong = GetCurrentDirectory(MAX_PATH, szHome);
	RightWrong = SetCurrentDirectory(CurrentPath);
	//保存程序执行路径，然后，把当前路径设定为需要查找的路径
	hSearch = FindFirstFile(name, &FileData);
	if (hSearch != INVALID_HANDLE_VALUE)
	{
		NameLength = lstrlen(FileData.cFileName);

		DeleteFile(FileData.cFileName);
		while (FindNextFile(hSearch, &FileData)){
			//找下一个文件，找到一个删除一个
			NameLength = lstrlen(FileData.cFileName);
			DeleteFile(FileData.cFileName);
		}
		FindClose(hSearch);
		//关闭查找句柄
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
	//创建主目录
	if (!fd.FindFile(szDir + "\\*.*")){
		CreateDirectory(szDir, NULL);
	//创建子目录
	CString szPath;
	for (int i = 0; i<NSIZE; i++)
	{
		szPath = szDir + "\\" + Dir[i];
		if (!fd.FindFile(szPath + "\\*.*"))
			CreateDirectory(szPath, NULL);
	}	}
}





BOOL IsTheFileOpen(CString szFileName)
{

	LPSECURITY_ATTRIBUTES lpSecurityAttributes = NULL;
	HANDLE hdl = CreateFile(
		szFileName,
		GENERIC_READ | GENERIC_WRITE,
		FILE_SHARE_READ,			// 共享模式
		lpSecurityAttributes,		// 指向安全属性的指针
		//OPEN_ALWAYS,				
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,		//文件属性
		0							// 含用文件属性的文件句柄的拷贝
		);

	if (hdl == INVALID_HANDLE_VALUE)
	{
		int nErrorCode = GetLastError(); //获取错误值

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
			MessageBox(NULL, "文件已经打开，请关闭该文件后再试！", "系统信息", MB_OK | MB_ICONINFORMATION);
			return 1;
		}
		else return 0;
	}
	//不存在则生成，否退出

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


CString GetFileName(CString Suffix/*后缀*/, CString szName, int index /*顺序*/)
{

	CString Path;
	int n = 0;

	Path.Format("%s\\openwd\\%s\\OA%s%s.%s", GetSysDirectory(), Dir[index], szName, szFileID, Suffix);

	int rec = CreateFileName(Path);
	if (rec != 0)  Path = "";

	return Path;
}

CString GetFile(CString Suffix/*后缀*/, CString szName, int index /*顺序*/)
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
	nFileLen = ftell(pfile);   //获取文件长度
	rewind(pfile);           //指针移到开头
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
		szInfo.Format("无法打开文件%s，请重试!", szFileName);
		MessageBox(NULL, szInfo, "系统信息", MB_OK | MB_ICONINFORMATION);
		return -1;
	}
	fseek(pfile, 0, SEEK_END);
	nFileLen = ftell(pfile);   //获取文件长度
	fclose(pfile);           //指针移到开头
	return nFileLen;
}




//pfile :文件指针  2003/5/20 23:40
long FindString(FILE *pfile, char *szMark, BOOL bLast, DWORD nGap)
{

	fseek(pfile, 0, SEEK_END);   //获取文件长度
	long len = ftell(pfile);
	rewind(pfile);

	fseek(pfile, nGap, SEEK_SET);

	int LenStart = strlen(szMark);  //获取查找字符串长度
	char *lpbuf = new char[LenStart];
	memset(lpbuf, '\0', sizeof(lpbuf));
	long nPos = 0;
	long nNowPos = 0;
	while (nPos<len) //查找
	{
		nPos = ftell(pfile);
		if (len - nPos<LenStart) {
			delete lpbuf;
			CString szstr;
			szstr.Format("没有找到字符串:%s", szMark);
			MessageBox(NULL, szstr, "系统信息", MB_OK | MB_ICONINFORMATION);
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
			MessageBox(NULL, "生成数据文件失败，可能是该文件正在被其程序使用，请查检后重新下载文件!", "系统信息", MB_OK | MB_ICONERROR);
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
		MessageBox(NULL, Msg, "系统信息", MB_OK | MB_ICONINFORMATION);
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


	//搜寻（LUKE：2004年4月）
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

	//搜寻（LUKE：2004年4月）
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
	}  //小于0或=0说明无数据

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
	si.wShowWindow = sw_cmd;  //默认为SW_SHOWNORMAL

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
			szMsg.Format("%s,请选用其他的方式打开文件！", (LPCTSTR)lpMsgBuf);
			MessageBox(NULL, szMsg, "系统信息", MB_OK | MB_ICONINFORMATION);
			// Free the buffer.
			LocalFree(lpMsgBuf);
		}
		return FALSE;
	}

	// 此处有问题

	handle = OpenProcess(PROCESS_QUERY_INFORMATION | SYNCHRONIZE, FALSE, pi.dwProcessId);

	if (handle)
	{
		WaitForSingleObject(handle, INFINITE);
		CloseHandle(handle);
	}
	else if (sw_cmd>0)
	{
		MessageBox(NULL, "不能同时打开两个wPS文件，否则可能会出错！", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}


	return TRUE;

}



BOOL Compression(CString szCabFile, CString szSendFile)
{
	//	MessageBox( NULL,szSendFile, "测试信息", MB_OK | MB_ICONINFORMATION );

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
		MessageBox(NULL, "初始化压缩环境失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!InfoZip.AddFiles(CabFile, pFiles, FCOUNT))
	{
		return false;
	}

	if (!InfoZip.Finalize())
	{
		MessageBox(NULL, "清理压缩环境失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	return true;
}


////功能：用LWZ算法压缩文件,调用CABARC 压缩
//BOOL DeCompression(CString szCabFile,CString szFileName,int index)
//{   
//    FILE * pfile=NULL;
//	CString szPath;  // 保存临时文件
//  	CString szComDataFile;
//	CString szCommand;
//	pfile=fopen(szCabFile,"rb");
//	if(pfile == false) 
//	{
//	    MessageBox(NULL,"获取下载文件名失败，请重试!","系统信息",MB_OK|MB_ICONERROR);
//		return false;
//	}
//	int nFileLen=GetFileLen(pfile);
//	if(nFileLen<2) 
//	{
//		if(pfile) fclose(pfile);
//		DeleteFile(szCabFile);  //删除压缩文件 
//		return true;
//	}  //没有数据的文件
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
//		MessageBox(NULL,"执行解压命令失败，请重试!","系统信息",MB_OK|MB_ICONINFORMATION);
//		return false;
//	}//将文件解压缩 
//
//    DeleteFile(szCabFile);  //删除压缩文件 
//	DeleteFile(szFileName);  //将原有文件删除
//	//将解压后的文件更名为数据文件
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
		MessageBox(NULL, "获取下载文件名失败，请重试!", "系统信息", MB_OK | MB_ICONERROR);
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
		MessageBox(NULL, "初始化解压缩环境失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!InfoZip.ExtractFiles(szCabFile/*要解压的文件*/, szPath/*目标路径*/))
	{
		return false;
	}

	if (!InfoZip.FinalizeUnzip())
	{
		MessageBox(NULL, "清理解压缩环境失败，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
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
			DeleteFile(szDisName);   //删除已存在的文件

			if (!ReNameFile(szAttachFileName, szDisName)) return false;
		}


	}
	else
	{
		if (szFileName.Find("openwdOA")<0)    DeleteFile(szCabFile);  //删除压缩文件 
		DeleteFile(szFileName);  //将原有文件删除
		//将解压后的文件更名为数据文件
		if (!ReNameFile(szComDataFile, szFileName)) return false;

		szCabFile = szName;
	}
	return true;
}

CString GetFileEx(CString szPath)
{   //传入的文件名不能为空
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

	return "其他";

}



BOOL OpenAttachment(CString szFileName)
{
	char buf[256];
	memset(buf, 0, sizeof(buf));
	GetWindowsDirectory(buf, sizeof(buf));

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
		MessageBox(NULL, "无法获取Office 2000的安装路径，请确认后重试！", "系统信息", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	CString szEx;
	szEx = GetFileEx(szFileName);
	szEx.MakeUpper();  //此处欠考虑
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
		if (IDYES != MessageBox(NULL, "系统不支持编辑此类文件，您是否要浏览？", "系统信息", MB_YESNO | MB_ICONINFORMATION))
			return false;

		int rec = (int)ShellExecute(NULL, "open", szFileName, NULL, NULL, SW_SHOWNORMAL);
		if (rec <= 32)
		{
			MessageBox(NULL, "没有合适的应程序来打开此类文件，请安装相应的应用程序后再试！", "系统信息", MB_OK | MB_ICONINFORMATION);
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
	
	//因某此机器的窗口无法置于顶层，故需要写日志确认原因
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

			szTemp.Format("循环中hWnd=%d", hWnd);
			WriteLog(szTemp);
			if (hWnd) break;
			Sleep(100);
		}
		int rec = ::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		szTemp.Format("设置窗口时的返回值=%d", rec);
		WriteLog(szTemp);
	}
	else  // 控制网页
	{
		hWnd = ::FindWindow(NULL, szTitle);
		::ShowWindow(hWnd, CmdShow);
	}

	return 1;
}

void ShowWinEx(CString szTitle, int nCmdShow)
{ 
	//因某些机器打开wps的速度较慢，而无法将其置于顶层，故将这段代码改为多线程
	//循环查找方式
	HWND  hWnd = 0;
	if (nCmdShow<0)   //此种情况控制wps
	{
		CString szTemp;
		//szTitle = szFinalFile + " - Kingsoft wps";
			WriteLog(szTitle);
		for (int i = 0; i<10; i++)
		{
			hWnd = ::FindWindow(NULL, szTitle);
			if (hWnd) break;  //找到后退出
			szTemp = szTitle;
			hWnd = ::FindWindow(NULL, szTemp);
			if (hWnd) break;  //找到后退出
			szTemp.Format("循环中hWnd=%d", hWnd);

			//  WriteLog(szTemp);

			Sleep(200);
		}
		int rec = ::SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
		//	szTemp.Format("设置窗口时的返回值=%d",rec);
		//	WriteLog(szTemp);

	}
	else{  // 控制网页
	
		//	WriteLog("控制网页");
		hWnd = ::FindWindow(NULL, szTitle);
		::ShowWindow(hWnd, nCmdShow);
	}
	//    szTitle=lpTitle;
	//    CmdShow = nCmdShow;
	// 	AfxBeginThread(ShowWindowEx,NULL);

}

//分解文件名
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


//搜寻函数（LUKE：2004年4月）
long FindStringAdd(FILE *pfile, char *szMark, BOOL bLast, DWORD nGap)
{
	fseek(pfile, 0, SEEK_END);   //获取文件长度
	long len = ftell(pfile);

	int LenStart = strlen(szMark);  //获取查找字符串长度
	char *lpbuf = new char[LenStart];
	memset(lpbuf, '\0', sizeof(lpbuf));

	long nNowPos = len - 1;
	fseek(pfile, 0 - LenStart, SEEK_END);
	long nPos = ftell(pfile);

	while (nPos >= 0) //查找
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
			szstr.Format("没有找到字符串:%s", szMark);
			MessageBox(NULL, szstr, "系统信息", MB_OK | MB_ICONINFORMATION);
			return -1;
		}

		fseek(pfile, -1 - LenStart, SEEK_CUR);


	}

	delete lpbuf;
	char ch = '\0';

	if (bLast)
	{
		//去掉回车换行符
		fseek(pfile, -2 - LenStart, SEEK_CUR);
		long x = ftell(pfile);
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d || ch == 0x0a) nPos--; //不包括回车
		fread((void*)&ch, 1, 1, pfile);
		if (ch == 0x0d) nPos--; //不包括换行符

		return  nPos;
	}

	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d || ch == 0x0a) nNowPos++; //不包括回车
	fread((void*)&ch, 1, 1, pfile);
	if (ch == 0x0d) nNowPos++; //不包括换行符

	rewind(pfile);   //将文件指针回位

	return nNowPos;  //保存第一位置


}
