/**
增加WPS支持
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


		//开始一个KingSoft Wps实例 
		wpsDoc::CApplication oWpsApp; 
		if (!oWpsApp.CreateDispatch("wps.Application")) 
		{ 
			MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
			return S_FALSE ; 
		} 

		oWpsApp.put_Visible(VARIANT_FALSE);   //显示Wps文档

		

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

	//开始一个Microsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 

	//建立一个新的文档 
	wpsDoc::CDocuments  oDocs;
	wpsDoc::CDocument0  oDoc;
	

	oDocs = oWpsApp.get_Documents();

	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
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

	//去掉菜单   
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 

	//建立一个新的文档 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),//文件名称
		covFalse,      //打开非wps文件时，是否进行转换
		covFalse,      //表示是否以只读方式打开文件
		covFalse,      //表示是否将打开的文档添加到“文件”菜单底部的最近使用过的文件列表中
		NULL,       //表示打开文档时所需要的密码
		NULL,       //如果打开的文件是模板类型，PasswordTemplate 参数表示打开模板时所需要的密码
		covFalse,      //当即将打开的文档是一个已经打开的文档时，需要用到此参数。参数为 True 时，表示放弃对已打开文档的所有尚未保存的修改，并将重新打开该文档；参数为 True 时，表示则直接激活已打开的文档
		NULL,       //表示文档修改之后，保存时所需要的密码
		NULL,       //如果打开的文件是模板类型，在模板修改之后，保存时所需要的密码
		0,          //表示打开文档时使用的文件转换器
		NULL,       //表示保存文档时的编码方式。
		covTrue,       //表示打开的文档是否显示在 WPS 应用程序中
		covFalse,      //表示是否修复打开的文档
	     0,         //表示文档中横排文字的排列方式
		 covFalse,      //表示在文字编码不能识别时，是否弹出“编码”对话框
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


BOOL wpsDoc::LastText(CString szTempleteFileName,/*被插入的文件名*/  CString szHeaderFileName/*文件名称*/,CString szDataFileName,CString szInfo)
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
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	wpsDoc::CDocuments oDocs; 
    wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();

	oWpsApp.put_Visible(VARIANT_FALSE);   //不显示Wps文档
	//打开正文主体进行接受痕迹（为了解决定稿时某写文件会丢失数据的问题2008、11、4zhanglt增加）
	oDoc.AttachDispatch(oDocs.Open(
		 //文件名称
		COleVariant(szTempleteFileName, VT_BSTR),
		covFalse,      //打开非wps文件时，是否进行转换
		covFalse,      //表示是否以只读方式打开文件
		covFalse,      //表示是否将打开的文档添加到“文件”菜单底部的最近使用过的文件列表中
		NULL,       //表示打开文档时所需要的密码
		NULL,       //如果打开的文件是模板类型，PasswordTemplate 参数表示打开模板时所需要的密码
		covFalse,      //当即将打开的文档是一个已经打开的文档时，需要用到此参数。参数为 True 时，表示放弃对已打开文档的所有尚未保存的修改，并将重新打开该文档；参数为 True 时，表示则直接激活已打开的文档
		NULL,       //表示文档修改之后，保存时所需要的密码
		NULL,       //如果打开的文件是模板类型，在模板修改之后，保存时所需要的密码
		0,          //表示打开文档时使用的文件转换器
		NULL,       //表示保存文档时的编码方式。
		covTrue,       //表示打开的文档是否显示在 WPS 应用程序中
		covFalse,      //表示是否修复打开的文档
		0,         //表示文档中横排文字的排列方式
		covFalse,      //表示在文字编码不能识别时，是否弹出“编码”对话框
		NULL)
		);

	oDoc.AcceptAllRevisions();
	oDoc.Save();



	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
	oDoc.AttachDispatch(oDocs.Open(
		//COleVariant(szFileName,VT_BSTR),  
		COleVariant(szTempleteFileName, VT_BSTR),
		covFalse,      //打开非wps文件时，是否进行转换
		covFalse,      //表示是否以只读方式打开文件
		covFalse,      //表示是否将打开的文档添加到“文件”菜单底部的最近使用过的文件列表中
		NULL,       //表示打开文档时所需要的密码
		NULL,       //如果打开的文件是模板类型，PasswordTemplate 参数表示打开模板时所需要的密码
		covFalse,      //当即将打开的文档是一个已经打开的文档时，需要用到此参数。参数为 True 时，表示放弃对已打开文档的所有尚未保存的修改，并将重新打开该文档；参数为 True 时，表示则直接激活已打开的文档
		NULL,       //表示文档修改之后，保存时所需要的密码
		NULL,       //如果打开的文件是模板类型，在模板修改之后，保存时所需要的密码
		0,          //表示打开文档时使用的文件转换器
		NULL,       //表示保存文档时的编码方式。
		covTrue,       //表示打开的文档是否显示在 WPS 应用程序中
		covFalse,      //表示是否修复打开的文档
		0,         //表示文档中横排文字的排列方式
		covFalse,    //表示在文字编码不能识别时，是否弹出“编码”对话框
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
	//		MessageBox(NULL,"没有发现正文的书签'BKbody'，请与系统管理员联系!","系统信息",MB_OK|MB_ICONINFORMATION);
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
	//		MessageBox(NULL,"模板没有发现书签'BKhead'，请与系统管理员联系!","系统信息",MB_OK|MB_ICONINFORMATION);
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
		//			szTemp.Format("模板没有发现书签%s，请与系统管理员联系!",szBookMark);
		//			MessageBox(NULL,szTemp,"系统信息",MB_OK|MB_ICONINFORMATION);
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
	oDoc.AcceptAllRevisions();   //接收参数
	oDoc.Save();
	oDoc.ReleaseDispatch();
	//WpsApp.Quit(vOpt, vOpt, vOpt);
	return true ;

}



BOOL wpsDoc::Stamp(CString szFileName,/*被插入的文件名*/ CString InserFileNames/*含有公章的文件名*/)
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	//建立一个新的文档 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
		covFalse,      //打开非wps文件时，是否进行转换
		covFalse,      //表示是否以只读方式打开文件
		covFalse,      //表示是否将打开的文档添加到“文件”菜单底部的最近使用过的文件列表中
		NULL,       //表示打开文档时所需要的密码
		NULL,       //如果打开的文件是模板类型，PasswordTemplate 参数表示打开模板时所需要的密码
		covFalse,      //当即将打开的文档是一个已经打开的文档时，需要用到此参数。参数为 True 时，表示放弃对已打开文档的所有尚未保存的修改，并将重新打开该文档；参数为 True 时，表示则直接激活已打开的文档
		NULL,       //表示文档修改之后，保存时所需要的密码
		NULL,       //如果打开的文件是模板类型，在模板修改之后，保存时所需要的密码
		0,          //表示打开文档时使用的文件转换器
		NULL,       //表示保存文档时的编码方式。
		covTrue,       //表示打开的文档是否显示在 WPS 应用程序中
		covFalse,      //表示是否修复打开的文档
		0,         //表示文档中横排文字的排列方式
		covFalse,      //表示在文字编码不能识别时，是否弹出“编码”对话框
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
		MessageBox(NULL,"模板没有发现加盖公章书签，请与网络中心联系!","系统信息",MB_OK|MB_ICONINFORMATION);
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

	inLinesp.Select();               //2003/7/11 修改
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


////浏览
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
//	//开始一个Kingsoft Wps实例 
//    wpsDoc::CApplication oWpsApp; 
//    if (!oWpsApp.CreateDispatch("Wps.Application")) 
//    { 
//        MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
//        return S_FALSE ; 
//    } 
//
//
//	//建立一个新的文档 
//    Documents oDocs; 
//    _Document oDoc;
//	oDocs = oWpsApp.get_Documents();
//	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 



	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
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
		//			szTemp.Format("模板没有发现书签%s，请与系统管理员联系!",szBookMark);
		//			MessageBox(NULL,szTemp,"系统信息",MB_OK|MB_ICONINFORMATION);
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

	oDoc.Save();   //保存文件
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return S_FALSE ; 
	} 



	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
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
		//			szTemp.Format("模板没有发现书签%s，请与系统管理员联系!",szBookMark);
		//			MessageBox(NULL,szTemp,"系统信息",MB_OK|MB_ICONINFORMATION);
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 

	MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	wpsDoc::CCommandBars0 mybars;
	wpsDoc::CCommandBar1  mybar;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
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
	//	{   //显示修改痕迹
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

	oDoc.Save();   //保存文件
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





	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return false ; 
	} 


	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_TRUE);   //显示Wps文档
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName,VT_BSTR),  
	
		covFalse,      //打开非wps文件时，是否进行转换
		covFalse,      //表示是否以只读方式打开文件
		covFalse,      //表示是否将打开的文档添加到“文件”菜单底部的最近使用过的文件列表中
		NULL,       //表示打开文档时所需要的密码
		NULL,       //如果打开的文件是模板类型，PasswordTemplate 参数表示打开模板时所需要的密码
		covFalse,      //当即将打开的文档是一个已经打开的文档时，需要用到此参数。参数为 True 时，表示放弃对已打开文档的所有尚未保存的修改，并将重新打开该文档；参数为 True 时，表示则直接激活已打开的文档
		NULL,       //表示文档修改之后，保存时所需要的密码
		NULL,       //如果打开的文件是模板类型，在模板修改之后，保存时所需要的密码
		0,          //表示打开文档时使用的文件转换器
		NULL,       //表示保存文档时的编码方式。
		covTrue,       //表示打开的文档是否显示在 WPS 应用程序中
		covFalse,      //表示是否修复打开的文档
		0,         //表示文档中横排文字的排列方式
		covFalse,      //表示在文字编码不能识别时，是否弹出“编码”对话框
		NULL
		)     
		);
	//以下代码为盖章

	//解除对文档的保护
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
		MessageBox(NULL,"封发时间标签丢失请跟管理员联系!","系统信息",MB_OK|MB_ICONINFORMATION);
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
		MessageBox(NULL,"模板没有发现加盖公章书签!","系统信息",MB_OK|MB_ICONINFORMATION);
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

	inLinesp.Select();    //2003/7/11 修改
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


	oDoc.Save();   //保存文件
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
	//开始一个Kingsoft Wps实例 
	wpsDoc::CApplication oWpsApp; 
	if (!oWpsApp.CreateDispatch("wps.Application")) 
	{ 
		MessageBox(NULL,"加保护时，创建Wps对象失败","系统信息",MB_OK | MB_SETFOREGROUND); 
		return  false; 
	} 


	wpsDoc::CDocuments oDocs; 
	wpsDoc::CDocument0 oDoc;
	oDocs = oWpsApp.get_Documents();
	oWpsApp.put_Visible(VARIANT_FALSE);   //显示Wps文档
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

	//MessageBox(NULL,szInfo,"GetWpsFileFromServer！头信息",MB_OK|MB_ICONINFORMATION);
	int index=1;  
	CString szTextFile;
	CString szPowerFile;
	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //下载文件

	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile=="") return false;
	szPowerFile= GetFileName("ini","P_",index);

	//下载完成，现在要进行打开文件的操作
	char fname[256];
	strcpy(fname,szTextFile);

	FILE * pf=NULL;
	pf=fopen(szPowerFile,"r");    
	if(pf==NULL) 
	{
		MessageBox(NULL,"获取权限出错,请重试！","系统信息",MB_OK|MB_ICONINFORMATION);
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
	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //下载文件

	szTextFile = GetFileName("wps","D_",index);
	if(szTextFile=="") return false;


	if(!FinalFaxTextWps(szTextFile,nPower)) {DeleteFile(GetIniName(index));return false;}
	WriteString("LastFileName",szTextFile,GetIniName(index));
	WriteString("IsNeedLoad","1",GetIniName(index));
	WriteString("Protect","1",GetIniName(index)); //将标志位置为保护状态

	//szFinalFile=GetFile("wps","D_",index);

	return true;
}
BOOL EditFaxEx(char * szInfo,char *szHeader,char * szUserName,int nPower,int bHaveTrace)
{
	int index=5;
	CString szFaxFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false;}  //下载文件

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

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //下载文件

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
	//MessageBox(NULL,szInfo,"系统信息szInfo",MB_OK|MB_ICONERROR);
	//MessageBox(NULL,szIniFile,"系统信息szIniFile",MB_OK|MB_ICONERROR);
	szSendFile=GetString("LastFileName",szIniFile);
	if(szSendFile=="") return false;

	if(!IsTheFileExist(szSendFile))
	{
		MessageBox(NULL,"要上传的文件不存在，请确认后再试！","系统信息",MB_OK|MB_ICONERROR);
		return false;
	}
	if(IsTheFileOpen(szSendFile))
	{
		MessageBox(NULL,"要上传的文件正在被应用程序使用，请关闭后再试！","系统信息",MB_OK|MB_ICONWARNING);
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
	if(!Compression(szCabFile,szFileName)) return false;   //如果压缩文件失败返回
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
			MessageBox(NULL,"打开上传文件出错，请重试!","系统信息",MB_OK|MB_ICONINFORMATION);
			return false;
		}
	}
	catch(CException *e)
	{
		char msg[400];
		memset(msg,0,sizeof(msg));
		e->GetErrorMessage(msg,sizeof(msg)-1);
		CString szMsg=msg;
		if(szMsg.Find("共享")>0)
		{
			MessageBox(NULL,"请关闭文档后再进行发送操作!","系统信息",MB_OK|MB_ICONSTOP);
		}
		else
		{
			MessageBox(NULL,msg,"系统信息",MB_OK|MB_ICONSTOP);
		}
		return false;
	}

	nFileLen=GetFileLen(pfile);   //获取文件长度
	if(nFileLen<1)
	{
		MessageBox(NULL,"要上传的是一个空文件，请重新下载后再试！","系统信息",MB_OK|MB_ICONINFORMATION);
		fclose(pfile);
		DeleteFile(szCabFile);
		DeleteFile(GetIniName(index));
		return false;
	}
	//此处可以添加发送文件的属性等

	int nInfoLen=strlen(szInfo);
	buf=new char[nFileLen+nInfoLen+1];
	memset(buf,0,sizeof(nFileLen+nInfoLen+1));

	strcpy(buf,szInfo);   

	int len =fread((void*)(buf+nInfoLen),1,nFileLen,pfile);

	if(len!=nFileLen)
	{
		MessageBox(NULL,"发送数据的长度不正确，请重新发送!","系统信息",MB_OK|MB_ICONERROR);
		if(pfile) fclose(pfile);
		delete buf;
		return false;
	}
	if(pfile) fclose(pfile);  


	if(!wpsDoc::WpsConnectionHttp(buf,nInfoLen+nFileLen,index,false/*表示发送数据*/))
	{
		delete buf;
		return false;
	}
	//删除目录
	//DeleteAll(index);
	DeleteDirFile(index);
	delete buf;
	return true;
}
BOOL wpsDoc::WpsConnectionHttp(char * TextBuf,DWORD nFileLen,int index,int bDownLoad,CString szAttachmentFileName)
{

	if(bDownLoad)   //>0 表示下载
	{
		//下载加解压缩工具
		if(!GetTheCabarcFile()) return false;

		int rec =wpsDoc::IsNeedLoad(index);
		if(rec==-1) return false;  //出错
		if(rec==0) return true;    //已经下载
	}

	CString Ip, Port, ServerURL;
	try
	{
		if(!GetIpAndPort(Ip,Port,ServerURL)) {  //获取端口、IP地址、及服务器名称
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

	FILE * pfile=NULL;      //保存服务器下载的信息
	CString szPath;  // 保存临时文件
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
			MessageBox(NULL,"服务器结束请求失败，请重试!","系统信息", MB_OK|MB_ICONINFORMATION);
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
				MessageBox(NULL,"无法生成临时下载文件，可能是网络正忙，请稍后重试!","系统信息",MB_OK|MB_ICONINFORMATION);
				return false;
			}
			DWORD AllCount=0;
			for(;;)    
			{
				int len = pHttpFile->Read(buf,sizeof(buf)-1); 
				AllCount +=len;
				if(len==0) break  ;							  //将服务器返回信息息全部读出
				fwrite((void*)buf,1,len,pfile);							 
				memset(buf,0,sizeof(buf));
			}   //保存文件结束 
			if(pfile) fclose(pfile);	
			CString szStr;
			szStr=buf;
			if(szStr=="large")
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL,"文件太大，无法进行编辑操作!","系统信息",MB_OK|MB_ICONINFORMATION);
				return false;
			}
			if(AllCount<100)
			{
				if ( pHttpFile  !=NULL)	delete pHttpFile;
				if ( pHttpServer!=NULL)	delete pHttpServer;
				INetSession.Close();

				MessageBox(NULL,"服务器没有返回信息，请稍后重试!","系统信息",MB_OK|MB_ICONINFORMATION);
				return false;

			}

		}
		else   
		{
			CString sztemp;

			bool issuccessed=false;
			int findposition=0;


			int len = pHttpFile->Read(buf,sizeof(buf)-1);    //从端口读取返回信息
			sztemp=buf;
			sztemp.MakeUpper();		

			//Luke(2004-05-10)
			while(findposition<len) //查找
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
				MessageBox(NULL,"上传文件失败，请重新提交!","系统信息",MB_OK|MB_ICONINFORMATION);
				return false;
			}
		}
		//释放内存空间
		if ( pHttpFile  !=NULL)	delete pHttpFile;
		if ( pHttpServer!=NULL)	delete pHttpServer;
		INetSession.Close();

	}
	catch (CInternetException *pInetEx)
	{   //释放内存空间
		char msg[400];
		memset(msg,0,sizeof(msg));
		pInetEx->GetErrorMessage(msg,sizeof(msg)-1);
		CString szError;
		szError.Format("%s请重试！",msg);
		MessageBox(NULL,szError,"系统信息",MB_OK|MB_ICONERROR);
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

	//判断下列文件是否打开,当文件名为空时，说明文件已经打开
	if(GetFileName("ini","P_",index)=="") return -1;  //权限
	if(GetFileName("wps","H_",index)=="") return -1;  //头文件
	if(GetFileName("wps","T_",index)=="") return -1;  //模板
	if(GetFileName("wps","D_",index)=="") return -1;  //数据文件
	if(GetFileName("bmp","B_",index)=="") return -1;  //公章

	//如果不清理文件，则每次都要下载
	if(AfxGetApp()->GetProfileString("Telecom","DeleteAllFile","")!="") return true;

	if(GetString("IsNeedLoad",GetIniName(index))=="1") return false;  //如果存在则不需要下载

	return true;
}

BOOL wpsDoc::MakeFile(CString szFileName,int index ,CString szAttachmentPath)
{
	FILE *pfile=NULL;
	pfile=fopen(szFileName,"rb");
	if(pfile==NULL)
	{
		MessageBox(NULL,"打开已下载的数据文件失败，请重试!","系统信息",MB_OK|MB_ICONINFORMATION);
		return false;
	}

	if( !SplitFile(pfile,GetFileName("ini","P_",index) ,"HEADSTART","HEADEND")         ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","H_",index) ,"FILEHEADSTART","FILEHEADEND") ) {fclose(pfile) ;return false;}
	if( !SplitFile(pfile,GetFileName("zip","T_",index) ,"TMPSTART","TMPEND")           ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","D_",index) ,"DATASTART","DATAEND")         ) {fclose(pfile); return false;}
	if( !SplitFile(pfile,GetFileName("zip","B_",index) ,"PICTURESTART","PICTUREEND")   ) {fclose(pfile); return false;}
	if(pfile) fclose(pfile);

	if(index==10)  //2003/11/26  添加了下载所有附件的功能
	{  //下载所有附件
		if(!DeCompression(GetFileName("zip","D_",index),szAttachmentPath,index)) return false;
	}
	else
	{
		//解压缩文件
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

	if(!wpsDoc::WpsConnectionHttp(szHeader,strlen(szHeader),index)){ return false; }  //下载文件
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
	else//如果无则从路径中取　
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
	{ //压缩之
		if(!Compression(szCabFile,szCommand)) return false;
		DeleteFile(szCommand);
	}
	else//不压缩，只改名
	{
		szCabFile=szCommand;
	}

	FILE * pfile=NULL;
	pfile=fopen(szCabFile,"rb");

	if(pfile==NULL) 
	{
		CString szInfo;
		szInfo.Format("无法打开文件%s，本次上传失败,请重试！",szFileName);
		MessageBox(NULL,szInfo,"系统信息",MB_OK|MB_ICONINFORMATION);
		return false;
	}
	DWORD nFileLen=0;
	fseek(pfile,0,SEEK_END);
	nFileLen=ftell(pfile);   //获取文件长度
	rewind(pfile);           //指针移到开头
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
		if(!wpsDoc::WpsConnectionHttp(buf,nFileLen+nlen,index,0)) //上传文件
		{
			delete buf;
			buf=NULL;
			return false;
		}
		delete buf;
	}
	else // 2003/7/9  上传大于10M的文件 
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

			//发送数据
			if(!wpsDoc::WpsConnectionHttp(buffer,nFileLen,index,0)) 
			{
				AfxGetApp()->WriteProfileString("Telecom","Large","0");
				if(pfile) fclose(pfile);
				free(buffer);
				return false;
			}
		}   //发送结束
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

	int index =9; //表示附件下载
	CString szInformation;
	szInformation.Format(szInfo,szFileName);
	strcpy(szInfo,szInformation);
	CString szAttachFile;

	if(!wpsDoc::WpsConnectionHttp(szInfo,strlen(szInfo),index)){return false; }  //下载文件

	//生成附件
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
		MessageBox(NULL,"制作附件副本出错！","系统信息",MB_OK|MB_ICONERROR);
		return false;
	}

	AfxGetApp()->DoWaitCursor(0);

	//编辑数据
	if(!OpenAttachment(szAttachFile)) {DeleteFile(GetIniName(index));return false;}

	CString szTempFileName ;
	CString sztemp=szFileName;

	//	sztemp.Replace(" ","");  
	szTempFileName.Format("%s\\openwd\\%s\\%s",GetSysDirectory(),Dir[index],sztemp);
	DeleteFile(szTempFileName);
	if(!ReNameFile(szAttachFile,szTempFileName)) return false;
	szAttachFile=szTempFileName;
	//发送数据
	WriteString("IsNeedLoad","1",GetIniName(index));     //将这些标志位写入，以便上传失败后再次打开
	WriteString("LastFileName",szAttachFile,GetIniName(index));

	//发送
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

	//清除原有数据
	CString szDownLoadPath;
	szDownLoadPath.Format("%s\\openwd\\%s",GetSysDirectory(),Dir[index]);
	DeleteDataFile(szDownLoadPath);

	if(szFileNames=="") 
	{
		MessageBox(NULL,"请选择要下载的附件名称再试!","系统信息",MB_OK|MB_ICONWARNING);
		return false;
	}

	//	SetIpAndPort("172.16.10.21",81,"servlet/ULoadBDoc");
	CString szInformation;
	//选择下载路径
	CBrowseDirDialog dlg;
	dlg.m_Title="选择下择路径";
	dlg.m_Path="";
	if(dlg.DoBrowse()==0) return 1;  //不下载

	CString szPath=dlg.m_Path;

	CStringArray szItem;
	CString szTempName;
	GetAllFileNames(szItem,szFileNames);
	int nCount=szItem.GetSize();  //获取要下载的文件数
	for(int i=0;i<nCount;i++)
	{
		szTempName=szItem[i];

		if(!JudgeFileIgnoreOrNot(szPath,szTempName)) continue;
		memset(InfoBuf,0,sizeof(InfoBuf));
		strcpy(InfoBuf,szInfo);
		szInformation.Format(InfoBuf,szItem[i]);
		memset(InfoBuf,0,sizeof(InfoBuf));
		strcpy(InfoBuf,szInformation);
		szA_Name=szItem[i];   //将文件名保起来以备下载后改名
		if(!wpsDoc::WpsConnectionHttp(InfoBuf,strlen(InfoBuf),index,1,szTempName)){ return false; }  //下载文件
	}
	return true;
}

BOOL wpsDoc::SendAttach(CString szInfo)
{

	int index =9;
	static char BASED_CODE szFilter[] ="所有文件(*.*)|*.*|WPS文件(*.WPS)|*.DOC|BMP文件(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
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
			else //发送
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
		MessageBox(NULL,szfile1+szfile2+"与系统文件重名，请改名后再发送，其余文件已发送成功！","系统信息",MB_OK|MB_ICONINFORMATION);
	}
	return true;
}
BOOL wpsDoc::SendMailEx(CString szInfo,float fPart /*以K为单位*/,float fTotal/*以兆为单位*/)
{
	fTotal*=1000;  

	int index =9;
	static char BASED_CODE szFilter[] ="所有文件(*.*)|*.*|WPS文件(*.WPS)|*.DOC|BMP文件(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
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
		for(;;)  //保存发送数据的名称
		{
			CString FileName = BrowseDialog.GetNextPathName(pos);
			DWORD nFileLen=GetFileLen(FileName);
			if(nFileLen<0) return false;   //读文件发生错误
			nAllSize+=nFileLen;  
			szItemNames.Add(FileName);
			if(pos==NULL) break;
		}

		float fAllSize=(float)nAllSize/1000;
		float fSize=(fTotal-fPart)/1000;  //转换为M

		if( fAllSize>(fTotal-fPart) ) 
		{
			szItemNames.RemoveAll();
			CString szText;
			szText.Format("总的附件大小为%.2f兆，您已经附加了%.2f兆，不能再附加超过%.2f兆的附件！",fTotal/1000,fPart/1000,fSize);
			MessageBox(NULL,szText,"系统信息",MB_OK|MB_ICONINFORMATION);
			return false;
		}  //判断结束，符合条件则发送数据

		//发送数据
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
			else //发送
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
		MessageBox(NULL,szfile1+szfile2+"与系统文件重名，请改名后再发送，其余文件已发送成功！","系统信息",MB_OK|MB_ICONINFORMATION);
	}
	return true;

}

