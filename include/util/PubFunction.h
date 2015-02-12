/*----------------------------------------------
功能：通用功能函数
------------------------------------------------*/
#include <string>
#include <cmath>
#include <stdio.h>
#include <windows.h>
using namespace std;

#define NSIZE 11
extern CString Dir[NSIZE];

extern CString szFileID;     //文件ID号
extern CString szTmpID;
extern CString szFinalFile;
extern CString szA_Name;


//功能：存储服务器IP及端口
BOOL SetIpAndPort(char * Ip/*IP地址*/, int  Port/*端口*/, char *ServerURL/*请求服务器URL*/, char *Password/*解锁密码*/);

//功能：获取服务器IP及端口
BOOL GetIpAndPort(char * Ip/*IP地址*/, int * Port/*端口*/, char *ServerURL/*请求服务器URL*/);

//功能：设置文件ID号
void SetID(char* FileID, char* TmpID = "");


BOOL GetTheCabarcFile();


//BOOL ConnectionHttp(char * TextBuf="",DWORD nFileLen=0,int index =1,int bDownLoad=1,CString szAttachmentFileName="");

long FindString(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);

long FindStringAdd(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);
//功能：从服务器下载借阅档案
//BOOL GetEmprstimoFileFromServer(char* szInfo,char* hide);

//功能：上传文件
//BOOL SendAttach(CString szPath);
//功能：上传邮件   2003/09/17
//BOOL SendMailEx(CString szInfo,float fPart,float fTotal);

string FillToEightBits(string sz);

void CleanPlaintextMark(int iPlaintextLength);

void SetFlagEx(int index, int nflag);

int GetFlagEx(int index);

void SetInSureFlagEx(int index, int nflag);

int  GetInSureFlagEx(int index);

void  DeleteFilesEx(int index);

void ShowWinEx(CString szTitle, int nCmdShow = SW_MINIMIZE);

BOOL ExecuteFile(CString cmd, CString szFileName, UINT sw_cmd = SW_SHOWNORMAL);

//BOOL MakeFile(CString szFileName,int index ,CString szAttachmentPath="");
int DownLoadAllAttachmentEx(char * szInfo, CString szFileNames);

CString GetFileName(CString Suffix/*后缀*/, CString szName, int index /*顺序*/);

CString GetIniName(int index);

CString GetFile(CString Suffix/*后缀*/, CString szName, int index /*顺序*/);
void CreateDir();
/************************************************************************/
/* 删除临时文件夹内的临时文件                                           */
/************************************************************************/
void DeleteDirFile(int index);

//BOOL IsTheFileExist(CString szFileName);
BOOL IsTheFileOpen(CString szFileName);
bool OnFileCopy(CString m_SrcName, CString m_DstName);
BOOL Compression(CString szCabFile, CString szSendFile);
DWORD GetFileLen(CString szFileName);
CString GetSysDirectory();
DWORD GetFileLen(CString szFileName);
DWORD GetFileLen(FILE * pfile);

//int  IsNeedLoad(int index);
BOOL IsTheFileExist(CString szFileName);

BOOL DeCompression(CString& szCabFile, CString szFileName, int index);
BOOL SplitFile(FILE *pfile, CString szFileName, char* szMarkStart, char* szMarkEnd);
BOOL OpenAttachment(CString szFileName);
BOOL ReNameFile(CString szSourceFile, CString szDesFile);
BOOL DeleteDataFile(CString szPath);
//分解文件名
void GetAllFileNames(CStringArray &szItems, CString szFileNames);
BOOL JudgeFileIgnoreOrNot(CString szPath, CString & szFileName);



