#include "StdAfx.h"
#include "include/word/word.h"
#include "include/word/msword.h"

#include "include/util/PubFunction.h"
#include "include/util/Regedit.h"

using namespace std;

BOOL wdocx::OpenWordFile(Word::_ApplicationPtr m_pWord, CString szFileName, CString szUserName, int nState, int bHaveTrace){
	//CoInitialize(NULL);


	COleVariant covTrue((short)TRUE),
				covFalse((short)FALSE),
				covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant varstrNull("");
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	Word::_DocumentPtr    m_pDoc;

	

	m_pDoc = m_pWord->Documents->Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		varstrNull,
		varstrNull,
		covFalse,
		varstrNull,
		varstrNull,
		vFormat,
		covFalse,
		covTrue,
		covFalse,
		covFalse,
		covFalse,
		varstrNull);
	m_pWord->Visible = VARIANT_TRUE;

	try{
		m_pDoc->put_TrackRevisions(VARIANT_FALSE);
	}
	catch (...){
		TRACE("Office 2013! \n");
	}
	switch (nState)
	{
	case  EDIT:

		if (m_pDoc->GetProtectionType() == 0 || m_pDoc->GetProtectionType() == 2){

			try{
				m_pDoc->Unprotect(COleVariant("Password"));
		
				
			}
			catch (...){
				TRACE("Office 2013!\n");
			}
		}
		try{
			m_pDoc->put_TrackRevisions(VARIANT_FALSE);
			m_pDoc->put_PrintRevisions(bHaveTrace);
			m_pDoc->put_ShowRevisions(bHaveTrace);
		}
		catch (...){
			TRACE("Office 2013!\n");
		}
		break;
	case  MODIFY:
		if (szUserName.GetLength() > 0) m_pWord->put_UserName(szUserName.AllocSysString());
		//This is used by word xp 
		if (m_pDoc->GetProtectionType() == 0 || m_pDoc->GetProtectionType() == 2){
			try{
				m_pDoc->Unprotect(COleVariant("Password"));
			}
			catch (...){
				TRACE("Office 2013!\n");
			}
		}

		try{
			m_pDoc->put_TrackRevisions(VARIANT_TRUE);
			m_pDoc->put_PrintRevisions(bHaveTrace);
			m_pDoc->put_ShowRevisions(bHaveTrace);
			//view.put_ShowInsertionsAndDeletions(bHaveTrace);
			m_pDoc->Protect(Word::wdAllowOnlyRevisions, covFalse, COleVariant("Password"), covOptional, covOptional);
			
		}
		catch (...){
			TRACE("Office 2013!\n");
		}
		break;
	case  READONLY:

		try{

			m_pDoc->PrintRevisions = bHaveTrace;
			m_pDoc->ShowRevisions = bHaveTrace;
			m_pDoc->Protect(Word::wdAllowOnlyFormFields, covFalse, COleVariant("Password"), covOptional, covOptional);
		}
		catch (...){
			TRACE("Office 2013!\n");
		}
		break;
	default:
		break;
	}

	/*

	oDoc.ReleaseDispatch();
	//	oWordApp.ReleaseDispatch();

	//	oWordApp.Quit(vOpt,vOpt,vOpt);
	*/


	//CoUninitialize();

	
	return true;
}

BOOL wdocx::GetDocFileFromServer(Word::_ApplicationPtr m_pWord, CString sFileID, CString szUserName, int nOpenMode, int bHaveTrace)
{


	CString szTextFile, szPowerFile;
	szTextFile.Format("%s\\openwd\\%s\\%s.%s", GetSysDirectory(), Dir[nOpenMode], sFileID, "doc");
	if (!DocConnDownHttp(sFileID, nOpenMode)){
		AfxMessageBox("文档下载（DocConnDownHttp）失败");
		return false; 
	}  //下载文件

	if (wdocx::OpenWordFile(m_pWord,szTextFile, szUserName, nOpenMode, bHaveTrace) == false) {
		DeleteFile(GetIniName(nOpenMode));
		AfxMessageBox("打开下载后的文件失败");
    	return false;
	}
	
	WriteString("LastFileName", szTextFile, GetIniName(nOpenMode));
	WriteString("IsNeedLoad", "1", GetIniName(nOpenMode));

	return true;
}

BOOL wdocx::SendDocFileToServer(char* szInfo, int nOpenMode)
{
	CString szIniFile = GetSysDirectory() + "\\openwd\\" + Dir[nOpenMode] + "\\" + szFileID + ".ini";

	CString szSendFile= GetString("LastFileName", szIniFile);

	if (szSendFile == ""){ 
		AfxMessageBox("发送的文件名称为空");
		return false; }

	if (!IsTheFileExist(szSendFile))
	{
		MessageBox(NULL, "要上传的文件不存在，请确认后再试！", "系统信息", MB_OK | MB_ICONERROR);
		return false;
	}

	if (IsTheFileOpen(szSendFile))
	{
		MessageBox(NULL, "要上传的文件正在被应用程序使用，请关闭后再试！", "系统信息", MB_OK | MB_ICONWARNING);
		return false;
	}
	/*
	
	if (GetString("Protect", GetIniName(nOpenMode)) == "1")
	{

		if (!SetPortect(szSendFile)) {

			return false;
		}
	}
	*/
	CString szFileName;

	if (szInfo[1] == '1')
		szFileName.Format("%s\\openwd\\%s\\%s_dg.doc", GetSysDirectory(), Dir[nOpenMode], szFileID);
	else
		szFileName.Format("%s\\openwd\\%s\\%s.doc", GetSysDirectory(), Dir[nOpenMode], szFileID);

	if (!OnFileCopy(szSendFile, szFileName)) { 
		AfxMessageBox("复制文件错误");
		return false; }

//	CString szCabFile;
//	szCabFile.Format("%s\\openwd\\%s\\TempDoc.zip", GetSysDirectory(), Dir[nOpenMode]);
	
	//if (!Compression(szCabFile, szFileName)){
	//	AfxMessageBox("压缩文件错误！");
	//	return false;   //如果压缩文件失败返回
	//}
	
	//szSendFile = szCabFile;
	
	DeleteFile(szFileName);

	//037165839926-15637102006


	FILE * pfile = NULL;
	int nFileLen = 0;

	char *buf = NULL;
	try
	{
		pfile = fopen(szSendFile, "rb");
	
		if (pfile == NULL)
		{
			MessageBox(NULL, "打开上传文件出错1111，请重试!", "系统信息", MB_OK | MB_ICONINFORMATION);
			return false;
		}
	}
	catch (CException *e)
	{
		char msg[400];
		memset(msg, 0, sizeof(msg));
		e->GetErrorMessage(msg, sizeof(msg)-1);
		CString szMsg = msg;
		if (szMsg.Find("共享")>0)
		{
			MessageBox(NULL, "请关闭文档后再进行发送操作!", "系统信息", MB_OK | MB_ICONSTOP);
		}
		else
		{
			MessageBox(NULL, msg, "系统信息", MB_OK | MB_ICONSTOP);
		}
		return false;
	}

	nFileLen = GetFileLen(pfile);   //获取文件长度
	

	
	if (nFileLen<1)
	{
		MessageBox(NULL, "要上传的是一个空文件，请重新下载后再试！", "系统信息", MB_OK | MB_ICONINFORMATION);
		fclose(pfile);
		pfile = NULL;
	
		//DeleteFile(szCabFile);
		DeleteFile(GetIniName(nOpenMode));
		return false;
	}
	//此处可以添加发送文件的属性等

	int nInfoLen = strlen(szInfo);

	buf = new char[nFileLen + nInfoLen + 1];
	
	memset(buf, 0, sizeof(nFileLen + nInfoLen + 1));

	strcpy(buf, szInfo);
	
	
	
	int len = fread((void*)(buf + nInfoLen), 1, nFileLen, pfile);

	//pfile = fopen("c:\\aa.zip", "wb");
	//len = fwrite((void*)(buf + nInfoLen), 1, nFileLen, pfile);
	
	if (len != nFileLen)
	{
		MessageBox(NULL, "发送数据的长度不正确，请重新发送!", "系统信息", MB_OK | MB_ICONERROR);
		if (pfile) { fclose(pfile); pfile = NULL;}
		delete buf;
		return false;
	}

	if (pfile){ fclose(pfile); pfile = NULL; }
	
	

	if (!DocConnUploadHttp(buf, nInfoLen + nFileLen, nOpenMode))
	{
		delete buf;
		return false;
	}


	//删除目录
	//DeleteAll(index);
	DeleteDirFile(nOpenMode);

	delete buf;


	return true;


}





