#ifndef OPENWD_WPS_H
#define OPENWD_WPS_H
/*----------------------------------------------
功能：存放所有操作Wps文件的功能函数
时间：2009、10、22
编写：zhanglt
------------------------------------------------*/
namespace wpsDoc {
#define EDIT   0    //编辑
#define MODIFY 1	//修改
#define READONLY   2	//浏览

	//打开文档的方式
#define wpsOpenFormatAuto        0   //自动方式打开 
#define wpsOpenFormatDocument    1   //文档方式打开 
#define wpsOpenFormatTemplate    2   //模板方式打开 
#define wpsOpenFormatRTF         3   //RTF 方式打开 
#define wpsOpenFormatText        4   //文本方式打开 
#define wpsOpenFormatUnicodeText 5   //Unicode 文本方式打开 
#define wpsOpenFormatWebPages    6   //网页方式打开 
	//文档保护类型
#define wpsAllowOnlyComments     1   //文档保护方式为修订
#define wpsAllowOnlyFormFields   2
#define wpsAllowOnlyReading      3   //文档档保存方式为只读
#define wpsAllowOnlyRevisions    0
#define wpsNoProtection          -1  //文档保存方式为不保护




	//功能：打开wps文档
	BOOL OpenWpsFile(char * szFileName/*wps文件名*/, char * szUserName, int nPower = 0/*权限*/, int bHaveTrace = 0);
	//功能：打开借阅文档
	BOOL OpenEmprstimoFile(char * szFileName, BOOL hide);

	//盖章
	BOOL Stamp(CString szFileName,/*被插入的文件名*/ CString InserFileNames/*含有公章的文件名*/);
	//定稿
	BOOL LastText(CString szTempleteFileName,/*被插入的文件名*/  CString szHeaderFileName/*文件名称*/, CString szDataFileName, CString szInfo);

	//以下为传真处理函数
	//浏览
	//BOOL LookUpWps(CString szFileName,int bHaveTrace);
	//正文处理
	BOOL EditFaxWps(CString szFileName, CString UserName, CString szHeader, int nPower, int bHaveTrace);
	//定稿  接收修改，并显示修改痕迹
	BOOL FinalFaxWps(CString szFileName, CString  szHeader);
	//正文定稿 排版
	BOOL FinalFaxTextWps(CString szFileName, int nPower);
	//盖章
	BOOL StampFaxWps(CString szFileName, CString szStampFile);
	//给文件加写保护
	BOOL SetPortect(CString szFileName);


	//功能：wps文档转为Html
	BOOL OpenHtmlFile(char * szFileName/*wps文件名*/, char * szUserName, int nPower = 0/*权限*/, int bHaveTrace = 0);
	//功能：从服务器下载Wps文件
	BOOL GetWpsFileFromServer(char* szInfo, char * szUsername = "", int bHaveTrace = 0);


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
	BOOL SendWpsFileToServer(char *szInfo = "", int index = 1);

	BOOL WpsConnectionHttp(char * TextBuf = "", DWORD nFileLen = 0, int index = 1, int bDownLoad = 1, CString szAttachmentFileName = "");
	int  IsNeedLoad(int index);
	BOOL MakeFile(CString szFileName, int index, CString szAttachmentPath);
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
