/**
����WPS֧��
**/
#include "StdAfx.h"
#include "wps.h"
#include "wps/kingsoftWPS.h"

#include "util/PubFunction.h"
#include "util/des.h"
#include "util/Regedit.h"
#include "util/BrowseDirDialog.h"
#include "afxdlgs.h"

BOOL  RestoreMenubar(BOOL hide)   
{   
	TRY{      
		COleVariant covTrue((short)TRUE), 
			covFalse((short)FALSE), 
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		COleVariant vTrue((short)TRUE), 
			vFalse((short)FALSE), 
			vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			vP( (short)true, VT_I2 ); 
		COleVariant vPP(short (1));
		COleVariant vMM(short (0));
		COleVariant vdSaveChanges(short(0));
		COleVariant vFormat(short(0));


		//��ʼһ��KingSoft Wpsʵ�� 
		wpsDoc::CApplication oWpsApp; 
		if (!oWpsApp.CreateDispatch("wps.Application")) 
		{ 
			MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
			return S_FALSE ; 
		} 

		oWpsApp.put_Visible(VARIANT_FALSE);   //��ʾWps�ĵ�

		

		wpsDoc::CCommandBars0 mybars;
		wpsDoc::CCommandBar1  mybar;

		mybars.AttachDispatch (oWpsApp.get_CommandBars(),TRUE);


		wpsDoc::CCommandBar1			menu;   
		wpsDoc::CCommandBarControls0    menus; 
		wpsDoc::CCommandBarPopup0       File,   Tools; 

		menu.AttachDispatch(mybars.get_ActiveMenuBar());   
		menus.AttachDispatch(menu.get_Controls());   

		File.AttachDispatch(menus.get_Item(COleVariant((short)1)),   TRUE);   
		Tools.AttachDispatch(menus.get_Item(COleVariant((short)6)), TRUE);

		File.put_Visible(true);   
		Tools.put_Visible(true);

		mybar.ReleaseDispatch();   
		mybars.ReleaseDispatch();

		CFrameWnd   *   pwnd=(CFrameWnd   *)AfxGetMainWnd();   
		pwnd->GetActiveFrame()->UpdateWindow   ();   


		oWpsApp.Quit(COleVariant((short)false), vOpt, vOpt);  
		oWpsApp.ReleaseDispatch();

	}   
	CATCH(CException,   e)   
	{   

		TCHAR   errormsg[255];   
		e->GetErrorMessage(errormsg,255,NULL);   
	}   
	END_CATCH   
		return   true;   
} 











BOOL HideMenubar(char * szFileName,BOOL hide)
{


	COleVariant covTrue((short)TRUE), 
		covFalse((short)FALSE), 
		covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vPP(short (1));
	COleVariant vMM(short (0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));

	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);

	//��ʼһ��Microsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 

	//����һ���µ��ĵ� 
	wpsDoc::CDocuments  oDocs;
	wpsDoc::CDocument0  oDoc;
	

	oDocs = oWpsApp.get_Documents();

	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		NULL,
		NULL,
		covFalse,
		NULL,
		NULL, 
		wpsOpenFormatAuto,
		0,
		covTrue,
		covFalse,
		NULL,
		NULL,
		NULL));

	wpsDoc::CCommandBars0   mybars;

	wpsDoc::CCommandBar1   mybar;
	//mybars.AttachDispatch (oDoc.GetCommandBars (),TRUE);

	mybar.AttachDispatch (mybars.get_Item (COleVariant(/*(short)1)*/"Standard")),TRUE);
	mybar.put_Visible(true);
	mybar.AttachDispatch(mybars.get_Item(COleVariant(/*(short)2*/"Formatting")), TRUE);
	mybar.put_Visible (true);

	//ȥ���˵�   
	wpsDoc::CCommandBar1    cmdBar;
	wpsDoc::CCommandBar1     menu;
	wpsDoc::CCommandBarControls0  menus;
	wpsDoc::CCommandBarPopup0 File, Tools;

	menu.AttachDispatch(mybars.get_ActiveMenuBar());   
	menus.AttachDispatch(menu.get_Controls());   

	File.AttachDispatch(menus.get_Item(COleVariant((short)1)),   TRUE);   
	Tools.AttachDispatch(menus.get_Item(COleVariant((short)6)),   TRUE);  

	File.put_Visible   (true);   
	Tools.put_Visible   (true);  




	/**
	File.AttachDispatch(menus.get_Item(COleVariant((short)1)),   TRUE);   
	Tools.AttachDispatch(menus.get_Item(COleVariant((short)6)),   TRUE);   
	File.put_Visible   (true);   
	Tools.put_Visible   (true);  
	**/

	oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);

	mybar.ReleaseDispatch();   
	mybars.ReleaseDispatch();
	oDoc.ReleaseDispatch();
	oWpsApp.ReleaseDispatch();

	oWpsApp.Quit(COleVariant((short)false), vOpt, vOpt);



	CMDIFrameWnd * pwnd=(CMDIFrameWnd *)AfxGetMainWnd();
	pwnd->GetActiveFrame ()->UpdateWindow ();


	return true;


}







BOOL wpsDoc::OpenWpsFile(char * szFileName,char * szUserName,int nPower,int bHaveTrace)

{
	//    if(FileIsOpen(szFileName)) return false;
	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vPP(short (1));
	COleVariant vMM(short (0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));

	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 

	//����һ���µ��ĵ� 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),//�ļ�����
		covFalse,      //�򿪷�wps�ļ�ʱ���Ƿ����ת��
		covFalse,      //��ʾ�Ƿ���ֻ����ʽ���ļ�
		covFalse,      //��ʾ�Ƿ񽫴򿪵��ĵ����ӵ����ļ����˵��ײ������ʹ�ù����ļ��б���
		NULL,       //��ʾ���ĵ�ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ�PasswordTemplate ������ʾ��ģ��ʱ����Ҫ������
		covFalse,      //�������򿪵��ĵ���һ���Ѿ��򿪵��ĵ�ʱ����Ҫ�õ��˲���������Ϊ True ʱ����ʾ�������Ѵ��ĵ���������δ������޸ģ��������´򿪸��ĵ�������Ϊ True ʱ����ʾ��ֱ�Ӽ����Ѵ򿪵��ĵ�
		NULL,       //��ʾ�ĵ��޸�֮�󣬱���ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ���ģ���޸�֮�󣬱���ʱ����Ҫ������
		0,          //��ʾ���ĵ�ʱʹ�õ��ļ�ת����
		NULL,       //��ʾ�����ĵ�ʱ�ı��뷽ʽ��
		covTrue,       //��ʾ�򿪵��ĵ��Ƿ���ʾ�� WPS Ӧ�ó�����
		covFalse,      //��ʾ�Ƿ��޸��򿪵��ĵ�
	     0,         //��ʾ�ĵ��к������ֵ����з�ʽ
		 covFalse,      //��ʾ�����ֱ��벻��ʶ��ʱ���Ƿ񵯳������롱�Ի���
		NULL)     
		);
	wpsDoc::CWindow0 win;
	win=oWpsApp.get_ActiveWindow();

	wpsDoc::CView0  view;
	view=win.get_View();

	oDoc.put_TrackRevisions(false);
	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));
	if(nPower==EDIT)
	{
		oDoc.put_TrackRevisions(false);  
		oDoc.put_PrintRevisions(bHaveTrace);  
		oDoc.put_ShowRevisions(bHaveTrace);
	}
	if(nPower==MODIFY) 
	{   
		oWpsApp.put_UserName(szUserName);   

		oDoc.put_TrackRevisions(true);  

		oDoc.put_PrintRevisions(bHaveTrace);  
		oDoc.put_ShowRevisions(bHaveTrace);

		try{
			view.put_ShowInsertionsAndDeletions(bHaveTrace);
		}
		catch(...){
			TRACE("Office 2000!\n");
		}

		oDoc.Protect(0,vFalse,COleVariant(Password),vFalse,vFalse);

	}
	else if(nPower==READONLY)
	{
		oDoc.put_PrintRevisions(bHaveTrace);
		oDoc.put_ShowRevisions(bHaveTrace);
		oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);
	}
	oDoc.ReleaseDispatch();

	return true;
}


BOOL wpsDoc::LastText(CString szTempleteFileName,/*��������ļ���*/  CString szHeaderFileName/*�ļ�����*/,CString szDataFileName,CString szInfo)
{

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vPP(short (1));
	COleVariant vMM(short (0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);

	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	wpsDoc::CDocuments oDocs; 
    wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();

	oWpsApp.put_Visible(VARIANT_FALSE);   //����ʾWps�ĵ�
	//������������н��ܺۼ���Ϊ�˽������ʱĳд�ļ��ᶪʧ���ݵ�����2008��11��4zhanglt���ӣ�
	oDoc.AttachDispatch(oDocs.Open(
		 //�ļ�����
		COleVariant(szTempleteFileName, VT_BSTR),
		covFalse,      //�򿪷�wps�ļ�ʱ���Ƿ����ת��
		covFalse,      //��ʾ�Ƿ���ֻ����ʽ���ļ�
		covFalse,      //��ʾ�Ƿ񽫴򿪵��ĵ����ӵ����ļ����˵��ײ������ʹ�ù����ļ��б���
		NULL,       //��ʾ���ĵ�ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ�PasswordTemplate ������ʾ��ģ��ʱ����Ҫ������
		covFalse,      //�������򿪵��ĵ���һ���Ѿ��򿪵��ĵ�ʱ����Ҫ�õ��˲���������Ϊ True ʱ����ʾ�������Ѵ��ĵ���������δ������޸ģ��������´򿪸��ĵ�������Ϊ True ʱ����ʾ��ֱ�Ӽ����Ѵ򿪵��ĵ�
		NULL,       //��ʾ�ĵ��޸�֮�󣬱���ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ���ģ���޸�֮�󣬱���ʱ����Ҫ������
		0,          //��ʾ���ĵ�ʱʹ�õ��ļ�ת����
		NULL,       //��ʾ�����ĵ�ʱ�ı��뷽ʽ��
		covTrue,       //��ʾ�򿪵��ĵ��Ƿ���ʾ�� WPS Ӧ�ó�����
		covFalse,      //��ʾ�Ƿ��޸��򿪵��ĵ�
		0,         //��ʾ�ĵ��к������ֵ����з�ʽ
		covFalse,      //��ʾ�����ֱ��벻��ʶ��ʱ���Ƿ񵯳������롱�Ի���
		NULL)
		);

	oDoc.AcceptAllRevisions();
	oDoc.Save();



	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		//COleVariant(szFileName,VT_BSTR),  
		COleVariant(szTempleteFileName, VT_BSTR),
		covFalse,      //�򿪷�wps�ļ�ʱ���Ƿ����ת��
		covFalse,      //��ʾ�Ƿ���ֻ����ʽ���ļ�
		covFalse,      //��ʾ�Ƿ񽫴򿪵��ĵ����ӵ����ļ����˵��ײ������ʹ�ù����ļ��б���
		NULL,       //��ʾ���ĵ�ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ�PasswordTemplate ������ʾ��ģ��ʱ����Ҫ������
		covFalse,      //�������򿪵��ĵ���һ���Ѿ��򿪵��ĵ�ʱ����Ҫ�õ��˲���������Ϊ True ʱ����ʾ�������Ѵ��ĵ���������δ������޸ģ��������´򿪸��ĵ�������Ϊ True ʱ����ʾ��ֱ�Ӽ����Ѵ򿪵��ĵ�
		NULL,       //��ʾ�ĵ��޸�֮�󣬱���ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ���ģ���޸�֮�󣬱���ʱ����Ҫ������
		0,          //��ʾ���ĵ�ʱʹ�õ��ļ�ת����
		NULL,       //��ʾ�����ĵ�ʱ�ı��뷽ʽ��
		covTrue,       //��ʾ�򿪵��ĵ��Ƿ���ʾ�� WPS Ӧ�ó�����
		covFalse,      //��ʾ�Ƿ��޸��򿪵��ĵ�
		0,         //��ʾ�ĵ��к������ֵ����з�ʽ
		covFalse,    //��ʾ�����ֱ��벻��ʶ��ʱ���Ƿ񵯳������롱�Ի���
		NULL
		)     
		);
		

	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));


	wpsDoc::CSelection sel;
	wpsDoc::CBookmark0 mark;
	wpsDoc::CBookmarks marks;
	wpsDoc::CRange rg;
	marks=oDoc.get_Bookmarks();
	int  rec=marks.Exists("BKbody");
	//	if(!rec) 
	//	{
	//		MessageBox(NULL,"û�з������ĵ���ǩ'BKbody'������ϵͳ����Ա��ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	//		return false;
	//	}
	if(rec)
	{
		mark=marks.Item(COleVariant("BKbody"));
		rg=mark.get_Range();
		rg.InsertFile(szDataFileName,COleVariant(""),covTrue,covFalse,covFalse);
	}





	marks=oDoc.get_Bookmarks();
	rec=marks.Exists("BKhead");

	//	if(!rec) 
	//	{
	//		MessageBox(NULL,"ģ��û�з�����ǩ'BKhead'������ϵͳ����Ա��ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	//		return false;
	//	}

	if(rec)
	{
		mark=marks.Item(COleVariant("BKhead"));
		rg=mark.get_Range();
		rg.InsertFile(szHeaderFileName,COleVariant(""),covTrue,covFalse,covFalse);
	}

	CString szBookMark;
	CString szValue;
	CString szTemp;
	for(;;)
	{

		int len=szInfo.Find("#|");
		if(len<=0) break;

		szTemp=szInfo.Left(len);
		szInfo=szInfo.Mid(len+2);

		len=szTemp.Find("&&");
		szBookMark=szTemp.Left(len);
		szValue=szTemp.Mid(len+2);

		rec=marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;
		//		}
		if(rec)
		{
			mark=marks.Item(COleVariant(szBookMark));
			rg=mark.get_Range();
			rg.Select();
			sel=oWpsApp.get_Selection();
			sel.TypeText(szValue);
		}
	}
	oDoc.AcceptAllRevisions();   //���ղ���
	oDoc.Save();
	oDoc.ReleaseDispatch();
	//WpsApp.Quit(vOpt, vOpt, vOpt);
	return true ;

}



BOOL wpsDoc::Stamp(CString szFileName,/*��������ļ���*/ CString InserFileNames/*���й��µ��ļ���*/)
{
	
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vPP(short (1));
	COleVariant vMM(short (0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	//����һ���µ��ĵ� 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
		covFalse,      //�򿪷�wps�ļ�ʱ���Ƿ����ת��
		covFalse,      //��ʾ�Ƿ���ֻ����ʽ���ļ�
		covFalse,      //��ʾ�Ƿ񽫴򿪵��ĵ����ӵ����ļ����˵��ײ������ʹ�ù����ļ��б���
		NULL,       //��ʾ���ĵ�ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ�PasswordTemplate ������ʾ��ģ��ʱ����Ҫ������
		covFalse,      //�������򿪵��ĵ���һ���Ѿ��򿪵��ĵ�ʱ����Ҫ�õ��˲���������Ϊ True ʱ����ʾ�������Ѵ��ĵ���������δ������޸ģ��������´򿪸��ĵ�������Ϊ True ʱ����ʾ��ֱ�Ӽ����Ѵ򿪵��ĵ�
		NULL,       //��ʾ�ĵ��޸�֮�󣬱���ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ���ģ���޸�֮�󣬱���ʱ����Ҫ������
		0,          //��ʾ���ĵ�ʱʹ�õ��ļ�ת����
		NULL,       //��ʾ�����ĵ�ʱ�ı��뷽ʽ��
		covTrue,       //��ʾ�򿪵��ĵ��Ƿ���ʾ�� WPS Ӧ�ó�����
		covFalse,      //��ʾ�Ƿ��޸��򿪵��ĵ�
		0,         //��ʾ�ĵ��к������ֵ����з�ʽ
		covFalse,      //��ʾ�����ֱ��벻��ʶ��ʱ���Ƿ񵯳������롱�Ի���
		NULL
		)     
		);

	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));

	oDoc.put_TrackRevisions(false);  


	wpsDoc::CBookmark0 mark;
	wpsDoc::CBookmarks marks;
	marks=oDoc.get_Bookmarks();

	int bStamp=0,bTime=0;
	bStamp=marks.Exists("BKgz");
	bTime=marks.Exists("BKregtime");

	if(bStamp==0 && bTime==0) 
	{
		MessageBox(NULL,"ģ��û�з��ּӸǹ�����ǩ����������������ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	if(bStamp) mark=marks.Item(COleVariant("BKgz"));
	else mark=marks.Item(COleVariant("BKregtime"));
	wpsDoc::CRange rg;
	rg=mark.get_Range();
	wpsDoc::CSelection sel;
	sel=oWpsApp.get_Selection();

	rg.Select();

	wpsDoc::CShapes0 shape;
	wpsDoc::CShape0 sp;
	shape=oDoc.get_Shapes();
	sel=oWpsApp.get_Selection();

	VARIANT vResult;
	vResult.vt=VT_DISPATCH;
	vResult.pdispVal =sel.get_Range(); 
	wpsDoc::CnlineShapes LineShapes;
	wpsDoc::CnlineShape  inLinesp;
	LineShapes=sel.get_InlineShapes();
	inLinesp = LineShapes.AddPicture(InserFileNames,covFalse,covTrue,&vResult);

	inLinesp.Select();               //2003/7/11 �޸�
	sp=	inLinesp.ConvertToShape();  
	sel=oWpsApp.get_Selection();

	wpsDoc::CShapeRange1 ShapeRg;
	
	wpsDoc::CWrapFormat  Format;

	ShapeRg=sel.get_ShapeRange();
	Format=ShapeRg.get_WrapFormat();
	

	ShapeRg.put_RelativeHorizontalPosition(3);
	ShapeRg.put_RelativeVerticalPosition(3);
	ShapeRg.put_Left(-999996);
	ShapeRg.put_Top(-999995);

	ShapeRg.ZOrder(3);
	Format.put_Type(5);

	// 	if(bStamp)
	//	{
	//	 	sp.IncrementLeft(-30);
	// 		sp.IncrementTop(90);
	//	}
	//	else
	//	{
	//		sp.IncrementLeft(16);
	// 	//	sp.IncrementTop(-50);
	//	}
	oDoc.ReleaseDispatch();
	oDoc.Save();

	//	oWpsApp.Quit(vOpt, vOpt, vOpt);

	return true;
}


////���
//BOOL LookUpWps(CString szFileName,int bHaveTrace)
//{
//	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
//	COleVariant vTrue((short)TRUE), 
//		vFalse((short)FALSE), 
//		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
//		vP( (short)true, VT_I2 ); 
//    COleVariant vPP(short (1));
//    COleVariant vMM(short (0));
//	COleVariant vdSaveChanges(short(0));
//	COleVariant vFormat(short(0));
//    
//	//��ʼһ��Kingsoft Wpsʵ�� 
//    wpsDoc::CApplication oWpsApp; 
//    if (!oWpsApp.CreateDispatch("Wps.Application")) 
//    { 
//        MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
//        return S_FALSE ; 
//    } 
//
//
//	//����һ���µ��ĵ� 
//    Documents oDocs; 
//    _Document oDoc;
//	oDocs = oWpsApp.get_Documents();
//	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
//	oDoc.AttachDispatch(oDocs.Open(
//		COleVariant(szFileName,VT_BSTR),  
//		covFalse,   
//		covFalse,   
//		covFalse,    
//		covOptional, 
//		covOptional, 
//		covFalse,    
//		covOptional, 
//		covOptional, 
//		covOptional, 
//		covOptional, 
//		covTrue)     
//							 );
//	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
//		oDoc.Unprotect(COleVariant("CNCOAzhonglu010"));
//
//	//oDoc.put_TrackRevisions(bHaveTrace);
//	 oDoc.put_ShowRevisions(bHaveTrace);
//
//	 oDoc.Protect(2,vFalse,COleVariant("CNCOAzhonglu010"));
// 	// oDoc.put_TrackRevisions(bHaveTrace);
//	 oDoc.ReleaseDispatch();
//	return true;
//}


BOOL EditFaxWps(CString szFileName,CString szUserName,CString szHeader,int nPower,int bHaveTrace)
{
	//    if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vWFT4(short (4));
	COleVariant vWFT2(short (2));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 



	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
		covFalse,   
		covFalse,
		covFalse,
		NULL, 
		NULL,
		covFalse,
		NULL,
		NULL,
		vWFT4,
		NULL,
		NULL,
		vWFT2,
		covTrue,
		covTrue,
		NULL
		)     
		);
	oDoc.put_TrackRevisions(false);
	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));

	//2003/7/11  
	wpsDoc::CWindow0 win;
	win=oWpsApp.get_ActiveWindow();

	wpsDoc::CView0  view;
	view=win.get_View();


	CString szBookMark;
	CString szValue;
	CString szTemp;
	wpsDoc::CSelection sel;
	wpsDoc::CBookmark0 mark;
	wpsDoc::CBookmarks marks;
	marks=oDoc.get_Bookmarks();
	wpsDoc::CRange rg;

	for(;;)
	{
		int len=szHeader.Find("#|");
		if(len<=0) break;

		szTemp=szHeader.Left(len);
		szHeader=szHeader.Mid(len+2);

		len=szTemp.Find("&&");
		szBookMark=szTemp.Left(len);
		szValue=szTemp.Mid(len+2);

		int rec=marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;
		//		}
		if(rec)
		{
			mark=marks.Item(COleVariant(szBookMark));
			rg=mark.get_Range();
			rg.put_End(rg.get_End()-1);
			rg.Select();
			sel=oWpsApp.get_Selection();
			sel.TypeText(szValue);
		}
	}

	if(nPower==EDIT)
	{
	
		oDoc.put_TrackRevisions(false);  
		oDoc.put_PrintRevisions(bHaveTrace);  
		oDoc.put_ShowRevisions(bHaveTrace);
	}
	else if(nPower==MODIFY) 
	{   
		oWpsApp.put_UserName(szUserName);   
		oDoc.put_TrackRevisions(true);  
		oDoc.put_PrintRevisions(bHaveTrace);  
		oDoc.put_ShowRevisions(bHaveTrace);

		try{view.put_ShowInsertionsAndDeletions(bHaveTrace);}
		catch(...){	TRACE("Office 2000!\n");}
		oDoc.Protect(0,vFalse,COleVariant(Password),vFalse,vFalse);

	}
	else if(nPower==READONLY)
	{
		oDoc.put_PrintRevisions(bHaveTrace);
		oDoc.put_ShowRevisions(bHaveTrace);
		oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);
	}

	oDoc.Save();   //�����ļ�
	oDoc.ReleaseDispatch();	

	return true;
}


BOOL FinalFaxWps(CString szFileName,CString  szHeader)
{
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vWFT4(short (4));
	COleVariant vWFT2(short (2));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 



	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
		covFalse,   
		covFalse,
		covFalse,
		NULL, 
		NULL,
		covFalse,
		NULL,
		NULL,
		vWFT4,
		NULL,
		NULL,
		vWFT2,
		covTrue,
		covTrue,
		NULL
		)     
		);
	oDoc.put_TrackRevisions(false);
	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));


	CString szBookMark;
	CString szValue;
	CString szTemp;

	wpsDoc::CSelection sel;
	wpsDoc::CBookmark0 mark;
	wpsDoc::CBookmarks marks;
	marks=oDoc.get_Bookmarks();
	wpsDoc::CRange rg;

	for(;;)
	{
		int len=szHeader.Find("#|");
		if(len<=0) break;

		szTemp=szHeader.Left(len);
		szHeader=szHeader.Mid(len+2);

		len=szTemp.Find("&&");
		szBookMark=szTemp.Left(len);
		szValue=szTemp.Mid(len+2);

		int rec=marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;		
		//		}
		if(rec)
		{
			mark=marks.Item(COleVariant(szBookMark));
			rg=mark.get_Range();
			rg.put_End(rg.get_End()-1);
			rg.Select();

			sel=oWpsApp.get_Selection();
			sel.TypeText(szValue);
		}
	}


	oDoc.put_TrackRevisions(false);  
	oDoc.put_PrintRevisions(false);  
	oDoc.put_ShowRevisions(false);


	oDoc.AcceptAllRevisions(); 

	return true;
}


BOOL FinalFaxTextWps(CString szFileName,int nPower)
{
	////	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vWFT4(short (4));
	COleVariant vWFT2(short (2));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 

	MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	wpsDoc::CCommandBars0 mybars;
	wpsDoc::CCommandBar1  mybar;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
		covFalse,   
		covFalse,
		covFalse,
		NULL, 
		NULL,
		covFalse,
		NULL,
		NULL,
		vWFT4,
		NULL,
		NULL,
		vWFT2,
		covTrue,
		covTrue,
		NULL
		)     
		);


//	mybars.AttachDispatch (oDoc.GetCommandBars(),TRUE);
	mybar.AttachDispatch (mybars.get_Item (COleVariant(/*(short)1)*/"Track Changes")),TRUE);
	mybar.put_Visible (false);
	mybar.put_Enabled(false);

	mybar.AttachDispatch (mybars.get_Item (COleVariant(/*(short)1)*/"Reviewing")),TRUE);
	mybar.put_Visible (false);
	mybar.put_Enabled(false);

	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));

	if(nPower==EDIT)
	{

		oDoc.put_TrackRevisions(false);  
		oDoc.put_PrintRevisions(false);  
		oDoc.put_ShowRevisions(false);

		oDoc.AcceptAllRevisions();
	}
	//	else if(nPower==MODIFY) 
	//	{   //��ʾ�޸ĺۼ�
	//		oDoc.Protect(0,vFalse,COleVariant("CNCOAzhonglu010"));
	//		oDoc.put_TrackRevisions(false);  
	//		oDoc.put_PrintRevisions(false);  
	//        oDoc.put_ShowRevisions(false);
	//	}
	else 
	{
		oDoc.put_PrintRevisions(false);
		oDoc.put_ShowRevisions(false);
		oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);
	}

	oDoc.Save();   //�����ļ�
	oDoc.ReleaseDispatch();

	return true;
}


BOOL StampFaxWps(CString szFileName,CString szStampFile)
{
	////	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vWFT4(short (4));
	COleVariant vWFT2(short(2));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//  CString   strDate;   
	// CTime   ttime   =   CTime::GetCurrentTime();   
	//strDate.Format("%d/%d/%d/%hh:%mm:%ss",ttime.GetYear(),ttime.GetMonth(),ttime.GetDay() );    

	//CTime t=CTime::GetCurrentTime(); 
	//TRACE(t.Format("%hh:%mm:%ss")); 
	COleDateTime oleDt=COleDateTime::GetCurrentTime();
	CString strDate=oleDt.Format("%Y/%m/%d/ %H:%M:%S");





	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"����Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //��ʾWps�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
	
		covFalse,      //�򿪷�wps�ļ�ʱ���Ƿ����ת��
		covFalse,      //��ʾ�Ƿ���ֻ����ʽ���ļ�
		covFalse,      //��ʾ�Ƿ񽫴򿪵��ĵ����ӵ����ļ����˵��ײ������ʹ�ù����ļ��б���
		NULL,       //��ʾ���ĵ�ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ�PasswordTemplate ������ʾ��ģ��ʱ����Ҫ������
		covFalse,      //�������򿪵��ĵ���һ���Ѿ��򿪵��ĵ�ʱ����Ҫ�õ��˲���������Ϊ True ʱ����ʾ�������Ѵ��ĵ���������δ������޸ģ��������´򿪸��ĵ�������Ϊ True ʱ����ʾ��ֱ�Ӽ����Ѵ򿪵��ĵ�
		NULL,       //��ʾ�ĵ��޸�֮�󣬱���ʱ����Ҫ������
		NULL,       //����򿪵��ļ���ģ�����ͣ���ģ���޸�֮�󣬱���ʱ����Ҫ������
		0,          //��ʾ���ĵ�ʱʹ�õ��ļ�ת����
		NULL,       //��ʾ�����ĵ�ʱ�ı��뷽ʽ��
		covTrue,       //��ʾ�򿪵��ĵ��Ƿ���ʾ�� WPS Ӧ�ó�����
		covFalse,      //��ʾ�Ƿ��޸��򿪵��ĵ�
		0,         //��ʾ�ĵ��к������ֵ����з�ʽ
		covFalse,      //��ʾ�����ֱ��벻��ʶ��ʱ���Ƿ񵯳������롱�Ի���
		NULL
		)     
		);
	//���´���Ϊ����

	//������ĵ��ı���
	if(oDoc.get_ProtectionType()==0||oDoc.get_ProtectionType()==2)
		oDoc.Unprotect(COleVariant(Password));

	oDoc.put_ShowRevisions(false);
	wpsDoc::CBookmark0 mark;
	wpsDoc::CBookmark0 bkprinttime;

	wpsDoc::CBookmarks marks;


	marks=oDoc.get_Bookmarks();


	int ibkprinttime;
	ibkprinttime=marks.Exists("bkprinttime");

	if( ibkprinttime==0) 
	{
		MessageBox(NULL,"�ⷢʱ���ǩ��ʧ�������Ա��ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	bkprinttime=marks.Item(COleVariant("bkprinttime"));
	wpsDoc::CRange rgbkprinttime;
	wpsDoc::CSelection selbkprinttime;

	rgbkprinttime=bkprinttime.get_Range();
	rgbkprinttime.Select();
	rgbkprinttime.put_Text("");
	//rg.SetText(strDate);   


	selbkprinttime=oWpsApp.get_Selection();
	//CFont font=selbkprinttime.GetFont();


	selbkprinttime.TypeText(strDate);


	//	oDoc.ReleaseDispatch();
	//	oDoc.Save();










	int bStamp=0,bTime=0;
	bStamp=marks.Exists("BKgz");

	if(bStamp==0 && bTime==0) 
	{
		MessageBox(NULL,"ģ��û�з��ּӸǹ�����ǩ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	if(bStamp) mark=marks.Item(COleVariant("BKgz"));
	wpsDoc::CRange rg;
	rg=mark.get_Range();

	rg.Select();
	
	wpsDoc::CShapes0   shape;
	wpsDoc::CShape0 sp;
	shape=oDoc.get_Shapes();

	wpsDoc::CSelection  sel;
	sel=oWpsApp.get_Selection();

	VARIANT vResult;
	vResult.vt=VT_DISPATCH;
	vResult.pdispVal =sel.get_Range(); 


	wpsDoc::CnlineShapes LineShapes;
	wpsDoc::CnlineShape  inLinesp;
	LineShapes=sel.get_InlineShapes();

	inLinesp = LineShapes.AddPicture(szStampFile,covFalse,covTrue,&vResult);

	inLinesp.Select();    //2003/7/11 �޸�
	sp=	inLinesp.ConvertToShape();  
	sel=oWpsApp.get_Selection();

	wpsDoc::CShapeRange1 ShapeRg;
	wpsDoc::CWrapFormat  Format;             

	ShapeRg=sel.get_ShapeRange();
	Format=ShapeRg.get_WrapFormat();

	ShapeRg.put_RelativeHorizontalPosition(3);
	ShapeRg.put_RelativeVerticalPosition(3);
	ShapeRg.put_Left(-999998);
	ShapeRg.put_Top(-999995);

	ShapeRg.ZOrder(4);
	Format.put_Type(5);


	//	sp.IncrementTop(-30);


	oDoc.Save();   //�����ļ�
	oDoc.ReleaseDispatch();
	return true;
}

BOOL SetPortect(CString szFileName)
{
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE), 
		vFalse((short)FALSE), 
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP( (short)true, VT_I2 ); 
	COleVariant vPP(short (4));
	COleVariant vMM(short (2));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password,0,sizeof(Password));
	GetUnlokPassword(Password);
	//��ʼһ��Kingsoft Wpsʵ�� 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"�ӱ���ʱ������Wps����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
		return  false; 
	} 


	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_FALSE);   //��ʾWps�ĵ�
	try
	{
		oDoc.AttachDispatch(oDocs.Open(
		   COleVariant(szFileName,VT_BSTR),  
			covFalse,   
			covFalse,
			covFalse,
			NULL, 
			NULL,
			covFalse,
			NULL,
			NULL,
			vPP,
			NULL,
			NULL,
			vMM,
			covTrue,
			covTrue,
			NULL
			)     
			);
	}
	catch(CException * e)
	{
		e->Delete();
		return false;
	}


	if(oDoc.get_ProtectionType()==2)  
		;
	else if(oDoc.get_ProtectionType()==0)
	{
		oDoc.Unprotect(COleVariant(Password));
		oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);
	}
	else {
		oDoc.Protect(2,vFalse,COleVariant(Password),vFalse,vFalse);
	}
	oDoc.Save();
	oWpsApp.Quit(vOpt, vOpt, vOpt);

	return true;

}

BOOL wpsDoc::GetWpsFileFromServer(char* szInfo,char * szUserName,int bHaveTrace)
{

	//MessageBox(NULL,szInfo,"GetWpsFileFromServer��ͷ��Ϣ",MB_OK|MB_ICONINFORMATION);
	int index=1;  
	CString szTextFile;
	CString szPowerFile;
	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //�����ļ�

	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile=="") return false;
	szPowerFile= GetFileName("ini","P_",index);

	//������ɣ�����Ҫ���д��ļ��Ĳ���
	char fname[256];
	strcpy(fname,szTextFile);

	FILE * pf=NULL;
	pf=fopen(szPowerFile,"r");    
	if(pf==NULL) 
	{
		MessageBox(NULL,"��ȡȨ�޳���,�����ԣ�","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	char buf[30];
	memset(buf,0,sizeof(buf));
	fgets(buf,sizeof(buf)-1,pf);
	if(pf) fclose(pf);
	int npower=atoi(buf);

	if(wpsDoc::OpenWpsFile(fname,szUserName,npower,bHaveTrace)==false) {DeleteFile(GetIniName(index)); return false;}
	WriteString("LastFileName",szTextFile,GetIniName(index));   
	WriteString("IsNeedLoad","1",GetIniName(index));         
	//szFinalFile=GetFile("wps","D_",index);
	return true;
}

BOOL StampFaxEx(char * szInfo)
{   
	int index =8;  //
	CString szFaxFile;
	CString szPicture;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){ return false; }  
	szFaxFile = GetFileName("wps","D_",index);
	if(szFaxFile=="") return false;
	szPicture = GetFileName("bmp","B_",index);
	if(szPicture=="") return false;

	if(!StampFaxWps(szFaxFile,szPicture)) {DeleteFile(GetIniName(index));return false;}

	WriteString("LastFileName",szFaxFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index));
	//szFinalFile=GetFile("wps","D_",index);

	return true;
}

BOOL FinalTextEx(char *szInfo,int nPower)
{
	int index = 3;
	CString szTextFile;
	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //�����ļ�

	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile=="") return false;


	if(!FinalFaxTextWps(szTextFile,nPower)) {DeleteFile(GetIniName(index));return false;}
	WriteString("LastFileName",szTextFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index)); //����־λ��Ϊ����״̬

	//szFinalFile=GetFile("wps","D_",index);

	return true;
}
BOOL EditFaxEx(char * szInfo,char *szHeader,char * szUserName,int nPower,int bHaveTrace)
{
	int index=5;
	CString szFaxFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false;}  //�����ļ�

	szFaxFile = GetFileName("wps","D_",index);
	if(szFaxFile=="") return false;


	if(!EditFaxWps(szFaxFile,szUserName,szHeader,nPower,bHaveTrace)) {DeleteFile(GetIniName(index));return false;}
	WriteString("LastFileName",szFaxFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	//szFinalFile=GetFile("doc","D_",index);

	return true;
}
BOOL FinalFaxEx(char *szInfo ,char * szHeader)
{ 
	int index=6;  
	CString szFaxFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //�����ļ�

	szFaxFile = GetFileName("wps","D_",index);
	if(szFaxFile=="") return false;

	if(!FinalFaxWps(szFaxFile,szHeader)) {DeleteFile(GetIniName(index));return false;}

	WriteString("LastFileName",szFaxFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index));
	//szFinalFile=GetFile("wps","D_",index);

	return true;
}
BOOL FinalFaxTextEx(char *szInfo,int nPower)
{
	int index = 7 ;
	CString szFaxFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){ return false; }  
	szFaxFile = GetFileName("wps","D_",index);
	if(szFaxFile=="") return false;

	if(!FinalFaxTextWps(szFaxFile,nPower)) {DeleteFile(GetIniName(index));return false;}


	WriteString("LastFileName",szFaxFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index));
	//szFinalFile=GetFile("wps","D_",index);

	return true;
}


BOOL SendWpsFileToServer(char* szInfo,int index)
{
	CString szSendFile;
	CString szIniFile=GetIniName(index);
	//MessageBox(NULL,szInfo,"ϵͳ��ϢszInfo",MB_OK|MB_ICONERROR);
	//MessageBox(NULL,szIniFile,"ϵͳ��ϢszIniFile",MB_OK|MB_ICONERROR);
	szSendFile=GetString("LastFileName",szIniFile);
	if(szSendFile=="") return false;

	if(!IsTheFileExist(szSendFile))
	{
		MessageBox(NULL,"Ҫ�ϴ����ļ������ڣ���ȷ�Ϻ����ԣ�","ϵͳ��Ϣ",MB_OK|MB_ICONERROR);
		return false;
	}
	if(IsTheFileOpen(szSendFile))
	{
		MessageBox(NULL,"Ҫ�ϴ����ļ����ڱ�Ӧ�ó���ʹ�ã���رպ����ԣ�","ϵͳ��Ϣ",MB_OK|MB_ICONWARNING);
		return false;
	}



	if(GetString("Protect",GetIniName(index))=="1")
	{   
		if(!SetPortect(szSendFile)) return false ;
	}  

	CString szFileName;


	if(szInfo[1]=='1')
		szFileName.Format("%s\\openwd\\%s\\%s_dg.wps",GetSysDirectory(),Dir[index],szFileID);
	else
		szFileName.Format("%s\\openwd\\%s\\%s.wps",GetSysDirectory(),Dir[index],szFileID);

	if(!OnFileCopy(szSendFile,szFileName)) return false;

	CString szCabFile;
	szCabFile.Format("%s\\openwd\\%s\\TempDoc.zip",GetSysDirectory(),Dir[index]);	
	if(!Compression(szCabFile,szFileName)) return false;   //���ѹ���ļ�ʧ�ܷ���
	szSendFile=szCabFile;

	DeleteFile(szFileName);

	FILE * pfile=NULL;
	int nFileLen=0;

	char *buf=NULL;
	try
	{
		pfile=fopen(szSendFile,"rb");
		if(pfile==NULL)
		{
			MessageBox(NULL,"���ϴ��ļ�������������!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
			return false;
		}
	}
	catch(CException *e)
	{
		char msg[400];
		memset(msg,0,sizeof(msg));
		e->GetErrorMessage(msg,sizeof(msg)-1);
		CString szMsg=msg;
		if(szMsg.Find("����")>0)
		{
			MessageBox(NULL,"��ر��ĵ����ٽ��з��Ͳ���!","ϵͳ��Ϣ",MB_OK|MB_ICONSTOP);
		}
		else
		{
			MessageBox(NULL,msg,"ϵͳ��Ϣ",MB_OK|MB_ICONSTOP);
		}
		return false;
	}

	nFileLen=GetFileLen(pfile);   //��ȡ�ļ�����
	if(nFileLen<1)
	{
		MessageBox(NULL,"Ҫ�ϴ�����һ�����ļ������������غ����ԣ�","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		fclose(pfile);
		DeleteFile(szCabFile);
		DeleteFile(GetIniName(index));
		return false;
	}
	//�˴��������ӷ����ļ������Ե�

	int nInfoLen=strlen(szInfo);
	buf=new char[nFileLen+nInfoLen+1];
	memset(buf,0,sizeof(nFileLen+nInfoLen+1));

	strcpy(buf,szInfo);   

	int len =fread((void*)(buf+nInfoLen),1,nFileLen,pfile);

	if(len!=nFileLen)
	{
		MessageBox(NULL,"�������ݵĳ��Ȳ���ȷ�������·���!","ϵͳ��Ϣ",MB_OK|MB_ICONERROR);
		if(pfile) fclose(pfile);
		delete buf;
		return false;
	}
	if(pfile) fclose(pfile);  


	if(!wpsDoc::WpsConnectionHttp(buf,nInfoLen+nFileLen,index,false/*��ʾ��������*/))
	{
		delete buf;
		return false;
	}
	//ɾ��Ŀ¼
	//DeleteAll(index);
	DeleteDirFile(index);
	delete buf;
	return true;
}
BOOL wpsDoc::WpsConnectionHttp(char * TextBuf,DWORD nFileLen,int index,int bDownLoad,CString szAttachmentFileName)
{

	if(bDownLoad)   //>0 ��ʾ����
	{
		//���ؼӽ�ѹ������
		if(!GetTheCabarcFile()) return false;

		int rec =wpsDoc::IsNeedLoad(index);
		if(rec==-1) return false;  //����
		if(rec==0) return true;    //�Ѿ�����
	}

	CString Ip, Port, ServerURL;
	try
	{
		if(!GetIpAndPort(Ip,Port,ServerURL)) {  //��ȡ�˿ڡ�IP��ַ��������������
			return false; 
		}

		/*
		if(AfxGetApp()->GetProfileString("Telecom","Large","")=="1")
		{
			ServerURL="servlet/ULoadBDoc";
		}*/
	}
	catch(CException * e)
	{
		e->ReportError();
		return false;
	}

	CInternetSession INetSession;
	CHttpConnection *pHttpServer=NULL;
	CHttpFile* pHttpFile=NULL;

	FILE * pfile=NULL;      //������������ص���Ϣ
	CString szPath;  // ������ʱ�ļ�
	szPath.Format("%s\\openwd\\%s\\TempDoc.dat",GetSysDirectory(),Dir[index]);	


	try 
	{   
		INetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT        ,30*60*1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_SEND_TIMEOUT		 ,30*60*1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_RECEIVE_TIMEOUT	 ,30*60*1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_SEND_TIMEOUT	 ,30*60*1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT,30*60*1000);

		INTERNET_PORT nport =atoi( Port);
		if(nport>0)
			pHttpServer= INetSession.GetHttpConnection(Ip,nport);
		else	
			pHttpServer= INetSession.GetHttpConnection(Ip);   



		pHttpFile= pHttpServer->OpenRequest(CHttpConnection::HTTP_VERB_POST, ServerURL, NULL, 1, 
			NULL, NULL, INTERNET_FLAG_DONT_CACHE);


		pHttpFile->SendRequestEx(nFileLen);  

		pHttpFile->Write(TextBuf,nFileLen);  
		if ( !(pHttpFile->EndRequest()) )    
		{
			MessageBox(NULL,"��������������ʧ�ܣ�������!","ϵͳ��Ϣ", MB_OK|MB_ICONINFORMATION);
			INetSession.Close();
			return false;
		}    


		char buf[1001];
		memset(buf,0,sizeof(buf));	
		if(bDownLoad)  
		{
			pfile=fopen(szPath,"wb");
			if(pfile==NULL)
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL,"�޷�������ʱ�����ļ���������������æ�����Ժ�����!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
				return false;
			}
			DWORD AllCount=0;
			for(;;)    
			{
				int len = pHttpFile->Read(buf,sizeof(buf)-1); 
				AllCount +=len;
				if(len==0) break  ;							  //��������������ϢϢȫ������
				fwrite((void*)buf,1,len,pfile);							 
				memset(buf,0,sizeof(buf));
			}   //�����ļ����� 
			if(pfile) fclose(pfile);	
			CString szStr;
			szStr=buf;
			if(szStr=="large")
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL,"�ļ�̫���޷����б༭����!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
				return false;
			}
			if(AllCount<100)
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();

				MessageBox(NULL,"������û�з�����Ϣ�����Ժ�����!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
				return false;

			}

		}
		else   
		{
			CString sztemp;

			bool issuccessed=false;
			int findposition=0;


			int len = pHttpFile->Read(buf,sizeof(buf)-1);    //�Ӷ˿ڶ�ȡ������Ϣ
			sztemp=buf;
			sztemp.MakeUpper();		

			//Luke(2004-05-10)
			while(findposition<len) //����
			{
				if(len-findposition<2) 
				{
					issuccessed=false;
					break;
				}

				int i=0;
				for(i;i<2;i++) 
				{
					int j;
					char tempmark[4]="OK";
					j=findposition+i;

					if(sztemp[j]!=tempmark[i]) 
						break;

				}


				if(i==2) {issuccessed=true; break;}

				findposition=findposition+1;
			}

			if(issuccessed==false) 
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL,"�ϴ��ļ�ʧ�ܣ��������ύ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
				return false;
			}
		}
		//�ͷ��ڴ�ռ�
		if ( pHttpFile  !=NULL)	delete pHttpFile;
		if ( pHttpServer!=NULL)	delete pHttpServer;
		INetSession.Close();

	}
	catch (CInternetException *pInetEx)
	{   //�ͷ��ڴ�ռ�
		char msg[400];
		memset(msg,0,sizeof(msg));
		pInetEx->GetErrorMessage(msg,sizeof(msg)-1);
		CString szError;
		szError.Format("%s�����ԣ�",msg);
		MessageBox(NULL,szError,"ϵͳ��Ϣ",MB_OK|MB_ICONERROR);
		pInetEx->Delete();
		if ( pHttpFile  !=NULL)	delete pHttpFile;
		if ( pHttpServer!=NULL)	delete pHttpServer;
		if ( pfile ) fclose(pfile);
		INetSession.Close();
		return false;
	}

	if(bDownLoad)  
	{	
		if(!wpsDoc::MakeFile(szPath,index,szAttachmentFileName)) return false;
	}

	return true;
}


int  wpsDoc::IsNeedLoad(int index)
{
	int nMark  = atoi(GetString("Mark",GetIniName(index)));
	int nInMark= atoi(GetString("Mark",GetIniName(index)));
	if(nMark+nInMark<=0) DeleteDirFile(index) ;    //DeleteAll(index);

	//�ж������ļ��Ƿ��,���ļ���Ϊ��ʱ��˵���ļ��Ѿ���
	if(GetFileName("ini","P_",index)=="") return -1;  //Ȩ��
	if(GetFileName("wps","H_",index)=="") return -1;  //ͷ�ļ�
	if(GetFileName("wps","T_",index)=="") return -1;  //ģ��
	if(GetFileName("wps","D_",index)=="") return -1;  //�����ļ�
	if(GetFileName("bmp","B_",index)=="") return -1;  //����

	//����������ļ�����ÿ�ζ�Ҫ����
	if(AfxGetApp()->GetProfileString("Telecom","DeleteAllFile","")!="") return true;

	if(GetString("IsNeedLoad",GetIniName(index))=="1") return false;  //�����������Ҫ����

	return true;
}

BOOL wpsDoc::MakeFile(CString szFileName,int index ,CString szAttachmentPath)
{
	FILE *pfile=NULL;
	pfile=fopen(szFileName,"rb");
	if(pfile==NULL)
	{
		MessageBox(NULL,"�������ص������ļ�ʧ�ܣ�������!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	if( !SplitFile(pfile,GetFileName("ini","P_",index) ,"HEADSTART","HEADEND")         ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","H_",index) ,"FILEHEADSTART","FILEHEADEND") ) {fclose(pfile) ;return false;}
	if( !SplitFile(pfile,GetFileName("zip","T_",index) ,"TMPSTART","TMPEND")           ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","D_",index) ,"DATASTART","DATAEND")         ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","B_",index) ,"PICTURESTART","PICTUREEND")   ) {fclose(pfile); return false;}
	if(pfile) fclose(pfile);

	if(index==10)  //2003/11/26  �������������и����Ĺ���
	{  //�������и���
		if(!DeCompression(GetFileName("zip","D_",index),szAttachmentPath,index)) return false;
	}
	else
	{
		//��ѹ���ļ�
		if(!DeCompression(GetFileName("zip","H_",index),GetFileName("wps","H_",index),index)) return false;
		if(!DeCompression(GetFileName("zip","T_",index),GetFileName("wps","T_",index),index)) return false;
		if(!DeCompression(GetFileName("zip","D_",index),GetFileName("wps","D_",index),index)) return false;
		if(!DeCompression(GetFileName("zip","B_",index),GetFileName("bmp","B_",index),index)) return false;
	}
	return true;
}

BOOL wpsDoc::InsuerDocument(char * szHeader,char * szSomeString)
{
	int index =2;  
	CString szTextFile;
	CString szTemFile;
	CString szHeadFile;

	if(!wpsDoc::WpsConnectionHttp(szHeader,strlen(szHeader),index)){ return false; }  

	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile  =="")  return false;
	szTemFile = GetFileName("wps","T_",index);
	if(szTextFile =="") return false;
	szHeadFile = GetFileName("wps","H_",index);
	if(szHeadFile =="") return false;

	if(!wpsDoc::LastText(szTemFile,szHeadFile,szTextFile,szSomeString)) {DeleteFile(GetIniName(index));return false;}

	WriteString("LastFileName",szTemFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index)); 

	szFinalFile=GetFile("wps","T_",index);


	return true;
}

BOOL wpsDoc::StampCover(char * szHeader)
{
   AfxMessageBox("aaaaaaaaaaaaaaaaaaaa");
	int index = 4;
	CString szTextFile;
	CString szPicture;

	if(!wpsDoc::WpsConnectionHttp(szHeader,strlen(szHeader),index)){ return false; }  //�����ļ�
	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile=="") return false;
	szPicture  = GetFileName("bmp","B_",index);
	if(szPicture=="")  return false;



    AfxMessageBox("bbbbbbbbbbbbbbbbbb");
	if(!wpsDoc::Stamp(szTextFile,szPicture)) {DeleteFile(GetIniName(index));return false;}
AfxMessageBox("ccccccccccccccccccccccccc");
	WriteString("LastFileName",szTextFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index));

	szFinalFile=GetFile("wps","D_",index);

	return true;

}



BOOL wpsDoc::SendData(CString szHeader, CString szFileName, int index)
{
	if(!GetTheCabarcFile()) return false;

	CString szPath =szFileName;
	CString szCabFile;
	CString szCommand;

	int len=szFileName.Find("#|");
	if(len>0)
	{
		szPath=szFileName.Left(len);
		szFileName=szFileName.Mid(len+2);   
	}
	else//��������·����ȡ��
	{
		char buffer[256];
		memset(buffer,0,sizeof(buffer));
		strcpy(buffer,szPath);
		szFileName="";
		for(int i=strlen(buffer)-1;i>=0;i--)
		{
			if(buffer[i]=='\\') break;
			szFileName=buffer[i]+szFileName;
		}
	}
	szCommand.Format("%s\\openwd\\%s\\openwdoa",GetSysDirectory(),Dir[index]);


	CString szTemp;	szTemp=szFileName; szTemp.MakeUpper();
	int nrec=szTemp.Find(".ZIP");
	if(!OnFileCopy(szPath,szCommand)) return false;
	szCabFile.Format("%s\\openwd\\%s\\TempDoc.zip",GetSysDirectory(),Dir[index]);	

	if(nrec<0)
	{ //ѹ��֮
		if(!Compression(szCabFile,szCommand)) return false;
		DeleteFile(szCommand);
	}
	else//��ѹ����ֻ����
	{
		szCabFile=szCommand;
	}

	FILE * pfile=NULL;
	pfile=fopen(szCabFile,"rb");

	if(pfile==NULL) 
	{
		CString szInfo;
		szInfo.Format("�޷����ļ�%s�������ϴ�ʧ��,�����ԣ�",szFileName);
		MessageBox(NULL,szInfo,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		return false;
	}
	DWORD nFileLen=0;
	fseek(pfile,0,SEEK_END);
	nFileLen=ftell(pfile);   //��ȡ�ļ�����
	rewind(pfile);           //ָ���Ƶ���ͷ
	if(nFileLen==0) {if(pfile) fclose(pfile); DeleteFile(szCabFile); return true;}

	CString szInfo ;//="f"+FileID+DBPath+"&^&%s#|#";

	szInfo.Format(szHeader,szFileName);
	int nlen=szInfo.GetLength();

	if(nFileLen<=10*1000*1000)
	{
		char * buf = new char [nFileLen+nlen];
		memset(buf,0,sizeof(buf));
		strcpy(buf,szInfo);
		fread((void*)(buf+nlen),1,nFileLen,pfile);
		if(pfile) fclose(pfile);
		if(!wpsDoc::WpsConnectionHttp(buf,nFileLen+nlen,index,0)) //�ϴ��ļ�
		{
			delete buf;
			buf=NULL;
			return false;
		}
		delete buf;
	}
	else // 2003/7/9  �ϴ�����10M���ļ� 
	{
		DWORD FS=5000000;
		int nindex = nFileLen/FS;
		bool bleave=0;
		if(nFileLen%FS){bleave=1; nindex++;}

		char *buffer= (char*)malloc(FS+nlen+100);

		AfxGetApp()->WriteProfileString("Telecom","Large","1");  


		DWORD nAllCount=0;
		nAllCount=nFileLen;
		for(int i=1;i<=nindex;i++)
		{
			CString szSequence,szLast;
			if(i<10)		szSequence.Format("00%d",i);
			else if(i<100)	szSequence.Format("0%d" ,i);
			else			szSequence.Format("%d"  ,i);
			szLast="#"+szSequence;
			memset(buffer,0,sizeof(buffer));
			strcpy(buffer,szInfo);


			if(i==nindex && bleave)  
			{
				fread((void*)(buffer+nlen),1,nAllCount,pfile);
				szLast+="y#";
				strcpy(buffer+nlen+nAllCount,szLast);
				nFileLen=nlen+nAllCount+szLast.GetLength();
			}
			else
			{
				fread((void*)(buffer+nlen),1,FS,pfile);
				nAllCount-=FS;
				szLast+="n#";
				strcpy(buffer+nlen+FS,szLast);
				nFileLen=nlen+FS+szLast.GetLength();
			}

			//��������
			if(!wpsDoc::WpsConnectionHttp(buffer,nFileLen,index,0)) 
			{
				AfxGetApp()->WriteProfileString("Telecom","Large","0");
				if(pfile) fclose(pfile);
				free(buffer);
				return false;
			}
		}   //���ͽ���
		if(pfile) fclose(pfile);
		free(buffer);
		AfxGetApp()->WriteProfileString("Telecom","Large","0");
	}

	DeleteFile(szCabFile);

	return true;
}


BOOL wpsDoc::DownLoad(char * szInfo,char * szUpInfo,char * szFileName)
{
	AfxGetApp()->DoWaitCursor(1);

	int index =9; //��ʾ��������
	CString szInformation;
	szInformation.Format(szInfo,szFileName);
	strcpy(szInfo,szInformation);
	CString szAttachFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //�����ļ�

	//���ɸ���
	char path[256];
	memset(path,0,sizeof(path));
	strcpy(path,szFileName);
	CString szEx;
	for(int i=strlen(path)-1;i>0;i--)
	{
		if(path[i]=='.') break;
		szEx=path[i]+szEx;
	}

	szAttachFile = GetFileName(szEx,"A_",index);
	if(szAttachFile=="") return false;

	if(!OnFileCopy(GetFileName("wps","D_",index),szAttachFile))
	{
		MessageBox(NULL,"������������������","ϵͳ��Ϣ",MB_OK|MB_ICONERROR);
		return false;
	}

	AfxGetApp()->DoWaitCursor(0);

	//�༭����
	if(!OpenAttachment(szAttachFile)) {DeleteFile(GetIniName(index));return false;}

	CString szTempFileName ;
	CString sztemp=szFileName;

	//	sztemp.Replace(" ","");  
	szTempFileName.Format("%s\\openwd\\%s\\%s",GetSysDirectory(),Dir[index],sztemp);
	DeleteFile(szTempFileName);
	if(!ReNameFile(szAttachFile,szTempFileName)) return false;
	szAttachFile=szTempFileName;
	//��������
	WriteString("IsNeedLoad","1",GetIniName(index));     //����Щ��־λд�룬�Ա��ϴ�ʧ�ܺ��ٴδ�
	WriteString("LastFileName",szAttachFile,GetIniName(index));

	//����
	sztemp=szAttachFile+"#|"+sztemp;

	AfxGetApp()->DoWaitCursor(1);

	if(!wpsDoc::SendData(szUpInfo,sztemp,index)) return false;
	DeleteFile(szAttachFile);

	//DeleteAll(index);
	DeleteDirFile(index);
	AfxGetApp()->DoWaitCursor(0);

	return true;
}


int wpsDoc::DownLoadAllAttachmentEx(char * szInfo,CString szFileNames)
{
	int index=10;
	char InfoBuf[256];
	//  memset(InfoBuf,0,sizeof(InfoBuf));

	//���ԭ������
	CString szDownLoadPath;
	szDownLoadPath.Format("%s\\openwd\\%s",GetSysDirectory(),Dir[index]);
	DeleteDataFile(szDownLoadPath);

	if(szFileNames=="") 
	{
		MessageBox(NULL,"��ѡ��Ҫ���صĸ�����������!","ϵͳ��Ϣ",MB_OK|MB_ICONWARNING);
		return false;
	}

	//	SetIpAndPort("172.16.10.21",81,"servlet/ULoadBDoc");
	CString szInformation;
	//ѡ������·��
	CBrowseDirDialog dlg;
	dlg.m_Title="ѡ������·��";
	dlg.m_Path="";
	if(dlg.DoBrowse()==0) return 1;  //������

	CString szPath=dlg.m_Path;

	CStringArray szItem;
	CString szTempName;
	GetAllFileNames(szItem,szFileNames);
	int nCount=szItem.GetSize();  //��ȡҪ���ص��ļ���
	for(int i=0;i<nCount;i++)
	{
		szTempName=szItem[i];

		if(!JudgeFileIgnoreOrNot(szPath,szTempName)) continue;
		memset(InfoBuf,0,sizeof(InfoBuf));
		strcpy(InfoBuf,szInfo);
		szInformation.Format(InfoBuf,szItem[i]);
		memset(InfoBuf,0,sizeof(InfoBuf));
		strcpy(InfoBuf,szInformation);
		szA_Name=szItem[i];   //���ļ����������Ա����غ����
		if(!wpsDoc::WpsConnectionHttp(InfoBuf,strlen(InfoBuf),index,1,szTempName)){ return false; }  //�����ļ�
	}
	return true;
}

BOOL wpsDoc::SendAttach(CString szInfo)
{

	int index =9;
	static char BASED_CODE szFilter[] ="�����ļ�(*.*)|*.*|WPS�ļ�(*.WPS)|*.DOC|BMP�ļ�(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
	CString szfile1="",szfile2="";
	char BufFileNames[25600];
	memset(BufFileNames,0,sizeof(BufFileNames));
	CFileDialog BrowseDialog(TRUE,"","",OFN_ALLOWMULTISELECT,szFilter,NULL);

	BrowseDialog.m_ofn.lpstrFile=BufFileNames;         //2003/8/23 22:31
	BrowseDialog.m_ofn.nMaxFile=sizeof(BufFileNames);  //2003/8/23 22:31

	int nres=BrowseDialog.DoModal();

	if(nres == IDOK)
	{

		int ncount =0;
		POSITION pos =BrowseDialog.GetStartPosition();
		CString file1=szFileID;
		CString file2=szFileID+"_dgc";


		AfxGetApp()->DoWaitCursor(1);

		for(;;)
		{
			CString FileName = BrowseDialog.GetNextPathName(pos);


			TRACE(FileName);TRACE("\n");
			if(FileName.Find(file1)>-1)
			{ 

				szfile1=file1;
			}
			else if(FileName.Find(file2)>-1)
			{
				szfile2=file2;

			}
			else //����
			{
				if(!wpsDoc::SendData(szInfo,FileName,index))   //2003/8/23 21:31
				{
					return false;
				}
			}
			if(pos==NULL) break;
			ncount++;
		}

	}

	if(szfile1!="" && szfile2!="")
	{
		MessageBox(NULL,szfile1+szfile2+"��ϵͳ�ļ���������������ٷ��ͣ������ļ��ѷ��ͳɹ���","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	}
	return true;
}
BOOL wpsDoc::SendMailEx(CString szInfo,float fPart /*��KΪ��λ*/,float fTotal/*����Ϊ��λ*/)
{
	fTotal*=1000;  

	int index =9;
	static char BASED_CODE szFilter[] ="�����ļ�(*.*)|*.*|WPS�ļ�(*.WPS)|*.DOC|BMP�ļ�(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
	CString szfile1="",szfile2="";
	char BufFileNames[25600];
	memset(BufFileNames,0,sizeof(BufFileNames));
	CFileDialog BrowseDialog(TRUE,"","",OFN_ALLOWMULTISELECT,szFilter,NULL);

	BrowseDialog.m_ofn.lpstrFile=BufFileNames;         //2003/8/23 22:31
	BrowseDialog.m_ofn.nMaxFile=sizeof(BufFileNames);  //2003/8/23 22:31
	CString file1=szFileID;
	CString file2=szFileID+"_dg";
	CStringArray  szItemNames;
	szItemNames.Add("test");
	szItemNames.RemoveAll();  

	int nres=BrowseDialog.DoModal();
	if(nres == IDOK)
	{
		POSITION pos =BrowseDialog.GetStartPosition();

		AfxGetApp()->DoWaitCursor(1);
		DWORD  nAllSize=0;
		for(;;)  //���淢�����ݵ�����
		{
			CString FileName = BrowseDialog.GetNextPathName(pos);
			DWORD nFileLen=GetFileLen(FileName);
			if(nFileLen<0) return false;   //���ļ���������
			nAllSize+=nFileLen;  
			szItemNames.Add(FileName);
			if(pos==NULL) break;
		}

		float fAllSize=(float)nAllSize/1000;
		float fSize=(fTotal-fPart)/1000;  //ת��ΪM

		if( fAllSize>(fTotal-fPart) ) 
		{
			szItemNames.RemoveAll();
			CString szText;
			szText.Format("�ܵĸ�����СΪ%.2f�ף����Ѿ�������%.2f�ף������ٸ��ӳ���%.2f�׵ĸ�����",fTotal/1000,fPart/1000,fSize);
			MessageBox(NULL,szText,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
			return false;
		}  //�жϽ���������������������

		//��������
		for(int i=0;i<szItemNames.GetSize();i++)
		{
			CString FileName=szItemNames.GetAt(i);
			if(FileName.Find(file1)>-1)
			{ 
				szfile1=file1;
			}
			else if(FileName.Find(file2)>-1)
			{
				szfile2=file2;
			}
			else //����
			{
				Sleep(100);
				if(!wpsDoc::SendData(szInfo,FileName,index))   //2003/8/23 21:31
				{
					return false;
				}
			}
		}
	}
	szItemNames.RemoveAll(); 

	if(szfile1!="" || szfile2!="")
	{
		MessageBox(NULL,szfile1+szfile2+"��ϵͳ�ļ���������������ٷ��ͣ������ļ��ѷ��ͳɹ���","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	}
	return true;

}
