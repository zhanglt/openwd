
/*----------------------------------------------
功能：word文件功能函数
------------------------------------------------*/
#ifndef OPENED_WORD_H
#define OPENED_WORD_H

namespace wdocx {

#define EDIT        0   //编辑状态
#define MODIFY      1   //修改状态
#define READONLY    2	//浏览状态
#define FINALEDIT   3	//定稿后编辑，自动接受痕迹
	//功能：打开word文档
	BOOL OpenWordFile(CString szFileName/*word文件名*/, CString szUserName, int nState = 0/*文件打开状态*/, int bHaveTrace = 0/*痕迹*/);
	//功能：打开借阅文档
	//BOOL OpenWordEmprstimoFile (char * szFileName,BOOL hide);
	//盖章
	BOOL Stamp(CString szFileName,/*被插入的文件名*/ CString InserFileNames/*含有公章的文件名*/);
	//定稿
	BOOL LastText(CString szTempleteFileName,/*被插入的文件名*/  CString szHeaderFileName/*文件名称*/, CString szDataFileName, CString szInfo);
	//以下为传真处理函数
	//浏览
	//BOOL LookUpWord(CString szFileName,int bHaveTrace);
	//正文处理
	BOOL EditFaxWord(CString szFileName, CString UserName, CString szHeader, int nPower, int bHaveTrace);
	//定稿  接收修改，并显示修改痕迹
	BOOL FinalFaxWord(CString szFileName, CString  szHeader);
	//正文定稿 排版
	BOOL FinalFaxTextWord(CString szFileName, int nPower);
	//盖章
	BOOL StampFaxWord(CString szFileName, CString szStampFile);
	//给文件加写保护
	BOOL SetPortect(CString szFileName);
	//功能：word文档转为Html
	BOOL OpenHtmlFile(char * szFileName/*word文件名*/, char * szUserName, int nPower = 0/*权限*/, int bHaveTrace = 0);
	//功能：从服务器下载Doc文件
	BOOL GetDocFileFromServer(CString szInfo, CString szUsername = "", int nState=0, int bHaveTrace = 0);
	//盖章
	BOOL StampFaxEx(char * szInfo);
	BOOL FinalTextEx(char *szInfo, int nPower);
	//正文处理
	BOOL EditFaxEx(char *szInfo, char * szHeader, char * UserName, int nPower, int bHaveTrace);
	//定稿
	BOOL FinalFaxEx(char * szInfo, char *szHeader);
	//正文定稿
	BOOL FinalFaxTextEx(char * szInfo, int nPower);
	//功能：向服务器上传文件
	BOOL SendDocFileToServer(char *szInfo = "", int index = 1);

	
	//功能：向服务器上传文件
	BOOL InsuerDocument(char * szHeader, char * szSomeString = "");
	BOOL StampCover(char * szHeader);
	BOOL SendData(CString szHeader, CString szFileName, int index);
	BOOL DownLoad(char * szInfo, char * szUpInfo, char * szFileName);
	int DownLoadAllAttachmentEx(char * szInfo, CString szFileNames);
	BOOL SendAttach(CString szInfo);
	BOOL SendMailEx(CString szInfo, float fPart /*以K为单位*/, float fTotal/*以兆为单位*/);
	}
#endif