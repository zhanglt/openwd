/*----------------------------------------------
���ܣ�ͨ�ù��ܺ���
------------------------------------------------*/
#include <string>
#include <cmath>
#include <stdio.h>
#include <windows.h>
using namespace std;

#define NSIZE 11
extern CString Dir[NSIZE];

extern CString szFileID;     //�ļ�ID��
extern CString szTmpID;
extern CString szFinalFile;
extern CString szA_Name;


//���ܣ��洢������IP���˿�
BOOL SetIpAndPort(char * Ip/*IP��ַ*/, int  Port/*�˿�*/, char *ServerURL/*���������URL*/, char *Password/*��������*/);

//���ܣ���ȡ������IP���˿�
BOOL GetIpAndPort(char * Ip/*IP��ַ*/, int * Port/*�˿�*/, char *ServerURL/*���������URL*/);

//���ܣ������ļ�ID��
void SetID(char* FileID, char* TmpID = "");


BOOL GetTheCabarcFile();


//BOOL ConnectionHttp(char * TextBuf="",DWORD nFileLen=0,int index =1,int bDownLoad=1,CString szAttachmentFileName="");

long FindString(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);

long FindStringAdd(FILE *pfile, char *szMark, BOOL bLast = 0, DWORD nGap = 0);
//���ܣ��ӷ��������ؽ��ĵ���
//BOOL GetEmprstimoFileFromServer(char* szInfo,char* hide);

//���ܣ��ϴ��ļ�
//BOOL SendAttach(CString szPath);
//���ܣ��ϴ��ʼ�   2003/09/17
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

CString GetFileName(CString Suffix/*��׺*/, CString szName, int index /*˳��*/);

CString GetIniName(int index);

CString GetFile(CString Suffix/*��׺*/, CString szName, int index /*˳��*/);
void CreateDir();
/************************************************************************/
/* ɾ����ʱ�ļ����ڵ���ʱ�ļ�                                           */
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
//�ֽ��ļ���
void GetAllFileNames(CStringArray &szItems, CString szFileNames);
BOOL JudgeFileIgnoreOrNot(CString szPath, CString & szFileName);


