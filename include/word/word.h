#include "include/word/WordEventSink.h"
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
	BOOL OpenWordFile(Word::_ApplicationPtr m_pWord, CString szFileName/*word文件名*/, CString szUserName, int nState = 0/*文件打开状态*/, int bHaveTrace = 0/*痕迹*/);
	//功能：从服务器下载Doc文件
	BOOL GetDocFileFromServer(Word::_ApplicationPtr m_pWord, CString sFileID, CString szUsername = "", int nOpenMode = 1, int bHaveTrace = 1);
	//功能：向服务器上传文件
	BOOL SendDocFileToServer(char *szInfo = "", int nOpenMode = 1);


	}
#endif