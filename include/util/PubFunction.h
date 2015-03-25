/*----------------------------------------------
功能：通用功能函数
------------------------------------------------*/
#ifndef OPENED_PUBFUNCTION_H
#define OPENED_PUBFUNCTION_H
#include <string>
#include <cmath>
#include <stdio.h>
#include <windows.h>
using namespace std;

#define NSIZE 7
extern CString Dir[NSIZE];

extern CString szFileID;     //文件ID号
extern CString szTmpID;
extern CString szFinalFile;
extern CString szA_Name;

BOOL DocConnDownHttp(CString  sFileID = "", int nOpenMode = 1, CString szAttachmentFileName = "");
BOOL DocConnUploadHttp(char *  TextBuf = "", DWORD nFileLen = 0, int index = 1, CString szAttachmentFileName = "");

int  IsNeedLoad(int index);
BOOL MakeFile(CString szFileName, int nOpenMode, CString szAttachmentPath);

BOOL DocConnectionHttp(CString TextBuf = "", DWORD nFileLen = 0, int index = 1, int bDownLoad = 1, CString szAttachmentFileName = "");
int  IsNeedLoad(int index);


//功能：获取服务器IP及端口
BOOL GetIpAndPort(CString &Ip/*IP地址*/, CString &Port/*端口*/, CString &ServerURL/*请求服务器URL*/);

void SetID(char* FileID, char* TmpID = "");

BOOL GetTheCabarcFile();

long FindString(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);

long FindStringAdd(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);

string FillToEightBits(string sz);

void CleanPlaintextMark(int iPlaintextLength);

void  DeleteFilesEx(int index);

void ShowWinEx(CString szTitle, int nCmdShow = SW_MINIMIZE);

BOOL ExecuteFile(CString cmd, CString szFileName, UINT sw_cmd = SW_SHOWNORMAL);

int DownLoadAllAttachmentEx(char * szInfo, CString szFileNames);

CString GetFileName(CString Suffix/*后缀*/, CString szName, int index /*顺序*/);

CString GetIniName(int index);

CString GetFile(CString Suffix/*后缀*/, CString szName, int index /*顺序*/);
void CreateDir();


//BOOL IsTheFileExist(CString szFileName);
BOOL IsTheFileOpen(CString szFileName);
BOOL IsTheFileExist(CString szFileName);

bool OnFileCopy(CString m_SrcName, CString m_DstName);
BOOL OpenAttachment(CString szFileName);

CString GetSysDirectory();
DWORD GetFileLen(CString szFileName);
DWORD GetFileLen(FILE * pfile);


BOOL SplitFile(FILE *pfile, CString szFileName, char* szMarkStart, char* szMarkEnd);
BOOL ReNameFile(CString szSourceFile, CString szDesFile);

BOOL DeleteDataFile(CString szPath);
void DeleteDirFile(int index);

//分解文件名
void GetAllFileNames(CStringArray &szItems, CString szFileNames);

#endif



