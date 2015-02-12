

#include "StdAfx.h"
#include "word/word.h"
#include "word/msword.h"
#include "util/PubFunction.h"
#include "util/Regedit.h"
#include "util/BrowseDirDialog.h"
//#include "afxdlgs.h"
//#include "des.h"
//#include "../../OpenWD.h"

BOOL wdocx::OpenWordFile(char * szFileName, char * szUserName, int nPower, int bHaveTrace)
{
	//    if(FileIsOpen(szFileName)) return false;
	COleVariant covTrue((short)TRUE),
		covFalse((short)FALSE),
		covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));


	//char Password[256];
	//memset(Password, 0, sizeof(Password));
	////GetUnlokPassword(Password);

	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;


	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);

		return S_FALSE;
	}

	//����һ���µ��ĵ� 
	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();

	 //��ʾWord�ĵ�
	oWordApp.put_Visible(VARIANT_TRUE);
		oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional
		)
		);

	//�����ļ������߲˵�
	/**
	wdocx::_CommandBars mybars;
	wdocx::CommandBar  mybar;
	mybars.AttachDispatch (oDoc.get_CommandBars (),TRUE);
	wdocx::CommandBar   menu;
	wdocx::CommandBarControls   menus;
	wdocx::CommandBarPopup   File,   Tools;
	menu.AttachDispatch(mybars.GetActiveMenuBar());
	menus.AttachDispatch(menu.GetControls());
	File.AttachDispatch(menus.GetItem(COleVariant((short)1)),   TRUE);
	Tools.AttachDispatch(menus.GetItem(COleVariant((short)6)),   TRUE);
	File.SetVisible (false);
	Tools.SetVisible (false);
	File.Reset(); //�˵���λ��һ��Ҫ��λ����Ȼ�ճ���wordҲ�������ļ������߲˵���
	Tools.Reset();
	mybar.ReleaseDispatch();
	mybars.ReleaseDispatch();
	**/
	//2003/7/11  
	wdocx::CWindow0  win;
	win = oWordApp.get_ActiveWindow();

	wdocx::CView0  view;
	view = win.get_View();

	try{ oDoc.put_TrackRevisions(false); }
	catch (...){ TRACE("Office 2000!\n"); }

	if (oDoc.get_ProtectionType()== 0 || oDoc.get_ProtectionType() == 2){
		try{ oDoc.Unprotect(COleVariant("Password")); }
		catch (...){ TRACE("Office 2000!\n"); }

	}

	if (nPower == EDIT)
	{
		try{
			oDoc.put_TrackRevisions(false);
			oDoc.put_PrintRevisions(bHaveTrace);
			oDoc.put_ShowRevisions(bHaveTrace);
		}
		catch (...){ 
			TRACE("Office 2013!\n"); 
		}

	}


	if (nPower == MODIFY)
	{
		oWordApp.put_UserName(szUserName);

		//This is used by word xp 


		try{
			oDoc.put_TrackRevisions(true);
			oDoc.put_PrintRevisions(bHaveTrace);
			oDoc.put_ShowRevisions(bHaveTrace);
		}
		catch (...){ TRACE("Office 2013!\n"); }




		try{ view.put_ShowInsertionsAndDeletions(bHaveTrace); }
		catch (...){ TRACE("Office 2013!\n"); }
		//  AfxMessageBox(Password);
		try{ oDoc.Protect(0, vFalse, COleVariant("Password"), covOptional, covOptional); }
		catch (...){ TRACE("Office 2013!\n"); }


	}
	else if (nPower == READONLY)
	{
		try{
			oDoc.put_PrintRevisions(bHaveTrace);
			oDoc.put_ShowRevisions(bHaveTrace);
			oDoc.Protect(2, vFalse, COleVariant("Password"), covOptional, covOptional);
		}
		catch (...){ TRACE("Office 2013!\n"); }

	}


	oDoc.ReleaseDispatch();
	//	oWordApp.ReleaseDispatch();

	//	oWordApp.Quit(vOpt,vOpt,vOpt);

	return true;
}


BOOL wdocx::LastText(CString szTempleteFileName,/*��������ļ���*/  CString szHeaderFileName/*�ļ�����*/, CString szDataFileName, CString szInfo)
{

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	////GetUnlokPassword(Password);

	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return false;
	}


	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();

	oWordApp.put_Visible(VARIANT_FALSE);   //����ʾWord�ĵ�
	//������������н��ܺۼ���Ϊ�˽������ʱĳд�ļ��ᶪʧ���ݵ�����2008��11��4���������ӣ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szDataFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional)
		);

	oDoc.AcceptAllRevisions();
	oDoc.Save();
	//oDoc.ReleaseDispatch();
	oWordApp.Quit(vOpt, vOpt, vOpt);
	//oDoc.Close(vTrue,vFormat,NULL);


	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return false;
	}

	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szTempleteFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional
		));


	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));


	wdocx::CSection sel;
	wdocx::CBookmark0 mark;
	wdocx::CBookmarks marks;
	wdocx::CRange rg;
	marks = oDoc.get_Bookmarks();
	int  rec = marks.Exists("BKbody");
	//	if(!rec) 
	//	{
	//		MessageBox(NULL,"û�з������ĵ���ǩ'BKbody'������ϵͳ����Ա��ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	//		return false;
	//	}
	if (rec)
	{
		mark = marks.Item(COleVariant("BKbody"));
		rg = mark.get_Range();
		rg.InsertFile(szDataFileName, COleVariant(""), covTrue, covFalse, covFalse);
	}





	marks = oDoc.get_Bookmarks();
	rec = marks.Exists("BKhead");

	//	if(!rec) 
	//	{
	//		MessageBox(NULL,"ģ��û�з�����ǩ'BKhead'������ϵͳ����Ա��ϵ!","ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
	//		return false;
	//	}

	if (rec)
	{
		mark = marks.Item(COleVariant("BKhead"));
		rg = mark.get_Range();
		rg.InsertFile(szHeaderFileName, COleVariant(""), covTrue, covFalse, covFalse);
	}


	CString szBookMark;
	CString szValue;
	CString szTemp;
	for (;;)
	{

		int len = szInfo.Find("#|");
		if (len <= 0) break;

		szTemp = szInfo.Left(len);
		szInfo = szInfo.Mid(len + 2);

		len = szTemp.Find("&&");
		szBookMark = szTemp.Left(len);
		szValue = szTemp.Mid(len + 2);

		rec = marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;
		//		}
		if (rec)
		{
			mark = marks.Item(COleVariant(szBookMark));
			rg = mark.get_Range();
			rg.Select();
			//sel = oWordApp.get_Selection();
			//sel.TypeText(szValue);
			
		}
	}
	oDoc.AcceptAllRevisions();   //���ղ���
	oDoc.Save();
	oDoc.ReleaseDispatch();
	//WordApp.Quit(vOpt, vOpt, vOpt);
	return true;

}



BOOL wdocx::Stamp(CString szFileName,/*��������ļ���*/ CString InserFileNames/*���й��µ��ļ���*/)
{
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);
	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return false;
	}

	//����һ���µ��ĵ� 
	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional)
		);


	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));

	oDoc.put_TrackRevisions(false);


	wdocx::CBookmark0 mark;
	wdocx::CBookmarks marks;
	marks = oDoc.get_Bookmarks();

	int bStamp = 0, bTime = 0;
	bStamp = marks.Exists("BKgz");
	bTime = marks.Exists("BKregtime");

	if (bStamp == 0 && bTime == 0)
	{
		MessageBox(NULL, "ģ��û�з��ּӸǹ�����ǩ����������������ϵ!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (bStamp) mark = marks.Item(COleVariant("BKgz"));
	else mark = marks.Item(COleVariant("BKregtime"));
	wdocx::CRange rg;
	rg = mark.get_Range();
	wdocx::CSelection sel;
	sel = oWordApp.get_Selection();
	rg.Select();

	wdocx::CShapes shape;
	wdocx::CShape sp;
	shape = oDoc.get_Shapes();
	sel = oWordApp.get_Selection();

	VARIANT vResult;
	vResult.vt = VT_DISPATCH;
	vResult.pdispVal = sel.get_Range();
	wdocx::CnlineShapes LineShapes;
	wdocx::CnlineShape  inLinesp;
	LineShapes = sel.get_InlineShapes();
	inLinesp = LineShapes.AddPicture(InserFileNames, covFalse, covTrue, &vResult);

	inLinesp.Select();               //2003/7/11 �޸�
	sp = inLinesp.ConvertToShape();
	sel = oWordApp.get_Selection();

	wdocx::CShapeRange ShapeRg;
	wdocx::CWrapFormat  Format;

	ShapeRg = sel.get_ShapeRange();
	Format = ShapeRg.get_WrapFormat();

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

	//	oWordApp.Quit(vOpt, vOpt, vOpt);

	return true;
}


////���
//BOOL LookUpWord(CString szFileName,int bHaveTrace)
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
//	//��ʼһ��Microsoft Wordʵ�� 
//    wdocx::CApplication oWordApp; 
//    if (!oWordApp.CreateDispatch("Word.Application")) 
//    { 
//        MessageBox(NULL,"����Word����ʧ��","ϵͳ��Ϣ",MB_OK | MB_SETFOREGROUND); 
//        return S_FALSE ; 
//    } 
//
//
//	//����һ���µ��ĵ� 
//    Documents oDocs; 
//    CDocument0 oDoc;
//	oDocs = oWordApp.get_Documents();
//	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
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


BOOL wdocx::EditFaxWord(CString szFileName, CString szUserName, CString szHeader, int nPower, int bHaveTrace)
{
	//    if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);
	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return S_FALSE;
	}



	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional)
		);
	oDoc.put_TrackRevisions(false);
	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));

	//2003/7/11  
	wdocx::CWindow0 win;
	win = oWordApp.get_ActiveWindow();

	wdocx::CView0  view;
	view = win.get_View();


	CString szBookMark;
	CString szValue;
	CString szTemp;
	wdocx::CSelection sel;
	wdocx::CBookmark0 mark;
	wdocx::CBookmarks marks;
	marks = oDoc.get_Bookmarks();
	wdocx::CRange rg;

	for (;;)
	{
		int len = szHeader.Find("#|");
		if (len <= 0) break;

		szTemp = szHeader.Left(len);
		szHeader = szHeader.Mid(len + 2);

		len = szTemp.Find("&&");
		szBookMark = szTemp.Left(len);
		szValue = szTemp.Mid(len + 2);

		int rec = marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;
		//		}
		if (rec)
		{
			mark = marks.Item(COleVariant(szBookMark));
			rg = mark.get_Range();
	
			rg.put_End(rg.get_End() - 1);
			rg.Select();
			sel = oWordApp.get_Selection();
			sel.TypeText(szValue);
		}
	}

	if (nPower == EDIT)
	{
		//oDoc.Protect(0,vFalse,COleVariant("CNCOAzhonglu010"));
		oDoc.put_TrackRevisions(false);
		oDoc.put_PrintRevisions(bHaveTrace);
		oDoc.put_ShowRevisions(bHaveTrace);
	}
	else if (nPower == MODIFY)
	{
		oWordApp.put_UserName(szUserName);
		oDoc.put_TrackRevisions(true);
		oDoc.put_PrintRevisions(bHaveTrace);
		oDoc.put_ShowRevisions(bHaveTrace);

		try{ view.put_ShowInsertionsAndDeletions(bHaveTrace); }
		catch (...){ TRACE("Office 2000!\n"); }
		oDoc.Protect(0, vFalse, COleVariant(Password),vFalse,vFalse);

	}
	else if (nPower == READONLY)
	{
		oDoc.put_PrintRevisions(bHaveTrace);
		oDoc.put_ShowRevisions(bHaveTrace);
		oDoc.Protect(2, vFalse, COleVariant(Password),vFalse,vFalse);
	}

	oDoc.Save();   //�����ļ�
	oDoc.ReleaseDispatch();

	return true;
}


BOOL wdocx::FinalFaxWord(CString szFileName, CString  szHeader)
{
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);
	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return S_FALSE;
	}



	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,    // Revert.
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,    // Revert.
		covOptional,
		covOptional,
		covOptional,
		covOptional)
		);
	oDoc.put_TrackRevisions(false);
	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));


	CString szBookMark;
	CString szValue;
	CString szTemp;

	wdocx::CSelection sel;
	wdocx::CBookmark0 mark;
	wdocx::CBookmarks marks;
	marks = oDoc.get_Bookmarks();
	wdocx::CRange rg;

	for (;;)
	{
		int len = szHeader.Find("#|");
		if (len <= 0) break;

		szTemp = szHeader.Left(len);
		szHeader = szHeader.Mid(len + 2);

		len = szTemp.Find("&&");
		szBookMark = szTemp.Left(len);
		szValue = szTemp.Mid(len + 2);

		int rec = marks.Exists(szBookMark);
		//		if(!rec) 
		//		{
		//			szTemp.Format("ģ��û�з�����ǩ%s������ϵͳ����Ա��ϵ!",szBookMark);
		//			MessageBox(NULL,szTemp,"ϵͳ��Ϣ",MB_OK|MB_ICONINFORMATION);
		//			return false;		
		//		}
		if (rec)
		{
			mark = marks.Item(COleVariant(szBookMark));
			rg = mark.get_Range();
			rg.put_End(rg.get_End() - 1);
			rg.Select();

			sel = oWordApp.get_Selection();
			sel.TypeText(szValue);
		}
	}


	oDoc.put_TrackRevisions(false);
	oDoc.put_PrintRevisions(false);
	oDoc.put_ShowRevisions(false);


	oDoc.AcceptAllRevisions();

	return true;
}


BOOL wdocx::FinalFaxTextWord(CString szFileName, int nPower)
{
	////	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);
	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return false;
	}

	MessageBox(NULL, "�����δ���Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;

	wdocx::CCommandBars mybars;
	wdocx::CCommandBar0  mybar;



	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional)     // �ɼ�
		);








	mybars.AttachDispatch(oDoc.get_CommandBars(), TRUE);
	mybar.AttachDispatch(mybars.get_Item(COleVariant(/*(short)1)*/"Track Changes")), TRUE);
	mybar.put_Visible(false);
	mybar.put_Enabled(false);

	mybar.AttachDispatch(mybars.get_Item(COleVariant(/*(short)1)*/"Reviewing")), TRUE);
	mybar.put_Visible(false);
	mybar.put_Enabled(false);

	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));

	if (nPower == EDIT)
	{
		//oDoc.Protect(0,vFalse,COleVariant("CNCOAzhonglu010"));

		try{
			oDoc.put_TrackRevisions(false);
			oDoc.put_PrintRevisions(false);
			oDoc.put_ShowRevisions(false);
			oDoc.AcceptAllRevisions();
		}
		catch (...){ TRACE("Office 2000!\n"); }


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
		try{
			oDoc.put_PrintRevisions(false);
			oDoc.put_ShowRevisions(false);
			oDoc.Protect(2, vFalse, COleVariant(Password),vFalse,vFalse);
		}
		catch (...){ TRACE("Office 2000!\n"); }
	}

	oDoc.Save();   //�����ļ�
	oDoc.ReleaseDispatch();

	return true;
}


BOOL wdocx::StampFaxWord(CString szFileName, CString szStampFile)
{
	////	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);
	//  CString   strDate;   
	// CTime   ttime   =   CTime::GetCurrentTime();   
	//strDate.Format("%d/%d/%d/%hh:%mm:%ss",ttime.GetYear(),ttime.GetMonth(),ttime.GetDay() );    

	//CTime t=CTime::GetCurrentTime(); 
	//TRACE(t.Format("%hh:%mm:%ss")); 
	COleDateTime oleDt = COleDateTime::GetCurrentTime();
	CString strDate = oleDt.Format("%Y/%m/%d/ %H:%M:%S");





	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "����Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return false;
	}


	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_TRUE);   //��ʾWord�ĵ�
	oDoc.AttachDispatch(oDocs.Open(
		COleVariant(szFileName, VT_BSTR),
		covFalse,
		covFalse,
		covFalse,
		covOptional,
		covOptional,
		covFalse,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covTrue,
		covOptional,
		covOptional,
		covOptional,
		covOptional)     // �ɼ�
		);


	//���´���Ϊ����

	//������ĵ��ı���
	if (oDoc.get_ProtectionType() == 0 || oDoc.get_ProtectionType() == 2)
		oDoc.Unprotect(COleVariant(Password));

	oDoc.put_ShowRevisions(false);
	wdocx::CBookmark0 mark;
	wdocx::CBookmark0 bkprinttime;

	wdocx::CBookmarks marks;


	marks = oDoc.get_Bookmarks();


	int ibkprinttime;
	ibkprinttime = marks.Exists("bkprinttime");

	if (ibkprinttime == 0)
	{
		MessageBox(NULL, "�ⷢʱ���ǩ��ʧ�������Ա��ϵ!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	bkprinttime = marks.Item(COleVariant("bkprinttime"));
	wdocx::CRange rgbkprinttime;
	wdocx::CSelection selbkprinttime;

	rgbkprinttime = bkprinttime.get_Range();
	rgbkprinttime.Select();
	rgbkprinttime.put_Text("");
	//rg.SetText(strDate);   


	selbkprinttime = oWordApp.get_Selection();
	//CFont font=selbkprinttime.GetFont();


	selbkprinttime.TypeText(strDate);


	//	oDoc.ReleaseDispatch();
	//	oDoc.Save();










	int bStamp = 0, bTime = 0;
	bStamp = marks.Exists("BKgz");

	if (bStamp == 0 && bTime == 0)
	{
		MessageBox(NULL, "ģ��û�з��ּӸǹ�����ǩ!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (bStamp) mark = marks.Item(COleVariant("BKgz"));
	wdocx::CRange rg;
	rg = mark.get_Range();

	rg.Select();
	wdocx::CShapes shape;
	wdocx::CShape sp;
	shape = oDoc.get_Shapes();

	wdocx::CSelection sel;
	sel = oWordApp.get_Selection();

	VARIANT vResult;
	vResult.vt = VT_DISPATCH;
	vResult.pdispVal = sel.get_Range();


	wdocx::CnlineShapes LineShapes;
	wdocx::CnlineShape  inLinesp;
	LineShapes = sel.get_InlineShapes();

	inLinesp = LineShapes.AddPicture(szStampFile, covFalse, covTrue, &vResult);

	inLinesp.Select();    //2003/7/11 �޸�
	sp = inLinesp.ConvertToShape();
	sel = oWordApp.get_Selection();

	wdocx::CShapeRange ShapeRg;
	wdocx::CWrapFormat  Format;

	ShapeRg = sel.get_ShapeRange();
	Format = ShapeRg.get_WrapFormat();

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

BOOL wdocx::SetPortect(CString szFileName)
{
	//	if(FileIsOpen(szFileName)) return false;

	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
		vP((short)true, VT_I2);
	COleVariant vPP(short(1));
	COleVariant vMM(short(0));
	COleVariant vdSaveChanges(short(0));
	COleVariant vFormat(short(0));
	char Password[256];
	memset(Password, 0, sizeof(Password));
	//GetUnlokPassword(Password);

	//AfxMessageBox(Password);

	//��ʼһ��Microsoft Wordʵ�� 
	wdocx::CApplication oWordApp;
	if (!oWordApp.CreateDispatch("Word.Application"))
	{
		MessageBox(NULL, "�ӱ���ʱ������Word����ʧ��", "ϵͳ��Ϣ", MB_OK | MB_SETFOREGROUND);
		return  false;
	}


	wdocx::CDocuments oDocs;
	wdocx::CDocument0 oDoc;
	oDocs = oWordApp.get_Documents();
	oWordApp.put_Visible(VARIANT_FALSE);   //��ʾWord�ĵ�
	try
	{
		oDoc.AttachDispatch(oDocs.Open(
			COleVariant(szFileName, VT_BSTR),
			covFalse,
			covFalse,
			covFalse,
			covOptional,
			covOptional,
			covFalse,    // Revert.
			covOptional,
			covOptional, // WritePasswordTemplate.
			covOptional,
			covOptional,
			covTrue,    // Revert.
			covOptional,
			covOptional, // 
			covOptional,
			covOptional)   
			);
	}
	catch (CException * e)
	{
		e->Delete();
		return false;
	}


	if (oDoc.get_ProtectionType() == 2)
		;
	else if (oDoc.get_ProtectionType() == 0)
	{
		AfxMessageBox(Password);
		oDoc.Unprotect(COleVariant(Password));
		oDoc.Protect(2, vFalse, COleVariant(Password),vFalse,vFalse);
	}
	else oDoc.Protect(2, vFalse, COleVariant(Password),vFalse,vFalse);

	oDoc.Save();
	oWordApp.Quit(vOpt, vOpt, vOpt);

	return true;

}


//************************************
// Method:    GetDocFileFromServer
// FullName:  wdocx::GetDocFileFromServer
// Access:    public 
// Returns:   BOOL
// Qualifier:
// Parameter: char * szInfo
// Parameter: char * szUserName
// Parameter: int bHaveTrace
//************************************
BOOL wdocx::GetDocFileFromServer(char* szInfo, char * szUserName, int bHaveTrace)
{

	//MessageBox(NULL,szInfo,"GetDocFileFromServer��ͷ��Ϣ",MB_OK|MB_ICONINFORMATION);
	int index = 1;
	CString szTextFile;
	CString szPowerFile;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }  //�����ļ�


	szTextFile = GetFileName("doc", "D_", index);
	if (szTextFile == "") return false;
	szPowerFile = GetFileName("ini", "P_", index);

	//������ɣ�����Ҫ���д��ļ��Ĳ���
	char fname[256];
	strcpy(fname, szTextFile);

	FILE * pf = NULL;
	pf = fopen(szPowerFile, "r");
	if (pf == NULL)
	{
		MessageBox(NULL, "��ȡȨ�޳���,�����ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	char buf[30];
	memset(buf, 0, sizeof(buf));
	fgets(buf, sizeof(buf)-1, pf);
	if (pf) fclose(pf);
	int npower = atoi(buf);


	/**
	int i=npower;
	char a[10];
	LPCSTR str;

	itoa(i, a, 10);
	str = a;

	MessageBox(NULL,str,"powerֵ",MB_OK|MB_ICONINFORMATION);

	**/
	if (wdocx::OpenWordFile(fname, szUserName, npower, bHaveTrace) == false) {
		DeleteFile(GetIniName(index));
		return false;
	}

	WriteString("LastFileName", szTextFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));


	//szFinalFile=GetFile("doc","D_",index);
	return true;
}

BOOL wdocx::StampFaxEx(char * szInfo)
{
	int index = 8;  //
	CString szFaxFile;
	CString szPicture;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }
	szFaxFile = GetFileName("doc", "D_", index);
	if (szFaxFile == "") return false;
	szPicture = GetFileName("bmp", "B_", index);
	if (szPicture == "") return false;

	if (!wdocx::StampFaxWord(szFaxFile, szPicture)) {
		DeleteFile(GetIniName(index));
		return false;
	}

	WriteString("LastFileName", szFaxFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	WriteString("Protect", "1", GetIniName(index));
	//szFinalFile=GetFile("doc","D_",index);

	return true;
}
BOOL wdocx::FinalTextEx(char *szInfo, int nPower)
{
	int index = 3;
	CString szTextFile;
	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }  //�����ļ�

	szTextFile = GetFileName("doc", "D_", index);
	if (szTextFile == "") return false;


	if (!wdocx::FinalFaxTextWord(szTextFile, nPower)) { DeleteFile(GetIniName(index)); return false; }
	WriteString("LastFileName", szTextFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	WriteString("Protect", "1", GetIniName(index)); //����־λ��Ϊ����״̬

	//szFinalFile=GetFile("doc","D_",index);

	return true;
}
BOOL wdocx::EditFaxEx(char * szInfo, char *szHeader, char * szUserName, int nPower, int bHaveTrace)
{
	int index = 5;
	CString szFaxFile;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }  //�����ļ�

	szFaxFile = GetFileName("doc", "D_", index);
	if (szFaxFile == "") return false;


	if (!wdocx::EditFaxWord(szFaxFile, szUserName, szHeader, nPower, bHaveTrace)) { DeleteFile(GetIniName(index)); return false; }
	WriteString("LastFileName", szFaxFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	//szFinalFile=GetFile("doc","D_",index);

	return true;
}
BOOL wdocx::FinalFaxEx(char *szInfo, char * szHeader)
{
	int index = 6;
	CString szFaxFile;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }  //�����ļ�

	szFaxFile = GetFileName("doc", "D_", index);
	if (szFaxFile == "") return false;

	if (!wdocx::FinalFaxWord(szFaxFile, szHeader)) { DeleteFile(GetIniName(index)); return false; }

	WriteString("LastFileName", szFaxFile, GetIniName(index));

	WriteString("IsNeedLoad", "1", GetIniName(index));

	WriteString("Protect", "1", GetIniName(index));

	//szFinalFile=GetFile("doc","D_",index);

	return true;
}
BOOL wdocx::FinalFaxTextEx(char *szInfo, int nPower)
{
	int index = 7;
	CString szFaxFile;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }
	szFaxFile = GetFileName("doc", "D_", index);
	if (szFaxFile == "") return false;

	if (!wdocx::FinalFaxTextWord(szFaxFile, nPower)) { DeleteFile(GetIniName(index)); return false; }


	WriteString("LastFileName", szFaxFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	WriteString("Protect", "1", GetIniName(index));
	//szFinalFile=GetFile("doc","D_",index);

	return true;
}

BOOL wdocx::SendDocFileToServer(char* szInfo, int index)
{
	CString szSendFile;
	CString szIniFile = GetIniName(index);
	//MessageBox(NULL,szInfo,"ϵͳ��ϢszInfo",MB_OK|MB_ICONERROR);
	//MessageBox(NULL,szIniFile,"ϵͳ��ϢszIniFile",MB_OK|MB_ICONERROR);
	szSendFile = GetString("LastFileName", szIniFile);
	if (szSendFile == "") return false;

	if (!IsTheFileExist(szSendFile))
	{
		MessageBox(NULL, "Ҫ�ϴ����ļ������ڣ���ȷ�Ϻ����ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		return false;
	}

	if (IsTheFileOpen(szSendFile))
	{
		MessageBox(NULL, "Ҫ�ϴ����ļ����ڱ�Ӧ�ó���ʹ�ã���رպ����ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONWARNING);
		return false;
	}



	if (GetString("Protect", GetIniName(index)) == "1")
	{

		if (!SetPortect(szSendFile)) {

			return false;
		}
	}



	CString szFileName;

	if (szInfo[1] == '1')
		szFileName.Format("%s\\unicom\\%s\\%s_dg.doc", GetSysDirectory(), Dir[index], szFileID);
	else
		szFileName.Format("%s\\unicom\\%s\\%s.doc", GetSysDirectory(), Dir[index], szFileID);

	if (!OnFileCopy(szSendFile, szFileName)) return false;

	CString szCabFile;
	szCabFile.Format("%s\\unicom\\%s\\TempDoc.zip", GetSysDirectory(), Dir[index]);

	//AfxMessageBox(szFileName);

	if (!Compression(szCabFile, szFileName)) return false;   //���ѹ���ļ�ʧ�ܷ���
	szSendFile = szCabFile;

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
			MessageBox(NULL, "���ϴ��ļ�������������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return false;
		}
	}
	catch (CException *e)
	{
		char msg[400];
		memset(msg, 0, sizeof(msg));
		e->GetErrorMessage(msg, sizeof(msg)-1);
		CString szMsg = msg;
		if (szMsg.Find("����")>0)
		{
			MessageBox(NULL, "��ر��ĵ����ٽ��з��Ͳ���!", "ϵͳ��Ϣ", MB_OK | MB_ICONSTOP);
		}
		else
		{
			MessageBox(NULL, msg, "ϵͳ��Ϣ", MB_OK | MB_ICONSTOP);
		}
		return false;
	}

	nFileLen = GetFileLen(pfile);   //��ȡ�ļ�����
	if (nFileLen<1)
	{
		MessageBox(NULL, "Ҫ�ϴ�����һ�����ļ������������غ����ԣ�", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		fclose(pfile);
		DeleteFile(szCabFile);
		DeleteFile(GetIniName(index));
		return false;
	}
	//�˴��������ӷ����ļ������Ե�

	int nInfoLen = strlen(szInfo);
	buf = new char[nFileLen + nInfoLen + 1];
	memset(buf, 0, sizeof(nFileLen + nInfoLen + 1));

	strcpy(buf, szInfo);

	int len = fread((void*)(buf + nInfoLen), 1, nFileLen, pfile);

	if (len != nFileLen)
	{
		MessageBox(NULL, "�������ݵĳ��Ȳ���ȷ�������·���!", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		if (pfile) fclose(pfile);
		delete buf;
		return false;
	}

	if (pfile) fclose(pfile);


	if (!wdocx::DocConnectionHttp(buf, nInfoLen + nFileLen, index, false/*��ʾ��������*/))
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


//************************************
// Method:    DocConnectionHttp
// FullName:  wdocx::DocConnectionHttp
// Access:    public 
// Returns:   BOOL
// Qualifier:
// Parameter: char * TextBuf
// Parameter: DWORD nFileLen
// Parameter: int index
// Parameter: int bDownLoad
// Parameter: CString szAttachmentFileName
//************************************
BOOL wdocx::DocConnectionHttp(char * TextBuf, DWORD nFileLen, int index, int bDownLoad, CString szAttachmentFileName)
{
	//MessageBox(NULL,TextBuf,"GetDocFileFromServer��!!!ͷ��Ϣ",MB_OK|MB_ICONINFORMATION);

	if (bDownLoad)   //>0 ��ʾ����
	{
		if (!GetTheCabarcFile()) return false; //���ؼӽ�ѹ������
		int rec = wdocx::IsNeedLoad(index);
		if (rec == -1) return false;             //����
		if (rec == 0) return true;               //�Ѿ�����
	}
	int	 Port = 0;//�������˿ں�
	char Ip[20];//������IP��ַ
	memset(Ip, 0, sizeof(Ip));
	char ServerURL[256];// �����URL
	memset(ServerURL, 0, sizeof(ServerURL));
	try
	{
		if (!GetIpAndPort(Ip, &Port, ServerURL)) {  //��ȡ�˿ڡ�IP��ַ��������������
			return false;
		}

		//2003/7/9 added by lhx
		if (AfxGetApp()->GetProfileString("Telecom", "Large", "") == "1")
		{
			memset(ServerURL, 0, sizeof(ServerURL));
			strcpy(ServerURL, "servlet/ULoadBDoc");
		}
	}
	catch (CException * e)
	{
		e->ReportError();
		return false;
	}

	CInternetSession INetSession;
	CHttpConnection *pHttpServer = NULL;
	CHttpFile* pHttpFile = NULL;

	FILE * pfile = NULL;      //������������ص���Ϣ
	CString szPath;         //������ʱ�ļ�
	szPath.Format("%s\\unicom\\%s\\TempDoc.dat", GetSysDirectory(), Dir[index]);
	try
	{
		INetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_SEND_TIMEOUT, 30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_DATA_RECEIVE_TIMEOUT, 30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_SEND_TIMEOUT, 30 * 60 * 1000);
		INetSession.SetOption(INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT, 30 * 60 * 1000);

		INTERNET_PORT nport = Port;

		if (nport>0)
			pHttpServer = INetSession.GetHttpConnection(Ip, nport);
		else
			pHttpServer = INetSession.GetHttpConnection(Ip);

		pHttpFile = pHttpServer->OpenRequest(CHttpConnection::HTTP_VERB_POST, ServerURL, NULL, 1, NULL, NULL, INTERNET_FLAG_DONT_CACHE);

		pHttpFile->SendRequestEx(nFileLen);

		pHttpFile->Write(TextBuf, nFileLen);


		if (!(pHttpFile->EndRequest()))
		{
			MessageBox(NULL, "��������������ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
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
				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "�޷�������ʱ�����ļ���������������æ�����Ժ�����!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
				return false;
			}
			DWORD AllCount = 0;
			for (;;)
			{
				int len = pHttpFile->Read(buf, sizeof(buf)-1);
				AllCount += len;
				if (len == 0) break;							  //��������������ϢϢȫ������
				fwrite((void*)buf, 1, len, pfile);
				memset(buf, 0, sizeof(buf));
			}   //�����ļ����� 
			if (pfile) fclose(pfile);
			CString szStr;
			szStr = buf;

			if (szStr == "large")
			{
				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "�ļ�̫���޷����б༭����!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
				return false;
			}

			if (AllCount<100)
			{
				if (pHttpFile != NULL)	delete pHttpFile;
				if (pHttpServer != NULL)	delete pHttpServer;
				INetSession.Close();
				MessageBox(NULL, "������û�з�����Ϣ�����Ժ�����!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
				return false;

			}

		}
		else
		{

			CString sztemp;

			bool issuccessed = false;
			int findposition = 0;

			int len = pHttpFile->Read(buf, sizeof(buf)-1);    //�Ӷ˿ڶ�ȡ������Ϣ
			sztemp = buf;
			sztemp.MakeUpper();

			//Luke(2004-05-10)
			while (findposition<len) //����
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
				MessageBox(NULL, "�ϴ��ļ�ʧ�ܣ��������ύ!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
				return false;
			}
		}
		//�ͷ��ڴ�ռ�
		if (pHttpFile != NULL)	delete pHttpFile;
		if (pHttpServer != NULL)	delete pHttpServer;
		INetSession.Close();
		//MessageBox(NULL,TextBuf,"44444444444",MB_OK|MB_ICONINFORMATION);
	}
	catch (CInternetException *pInetEx)
	{   //�ͷ��ڴ�ռ�
		char msg[400];
		memset(msg, 0, sizeof(msg));
		pInetEx->GetErrorMessage(msg, sizeof(msg)-1);
		CString szError;
		szError.Format("%s�����ԣ�", msg);
		MessageBox(NULL, szError, "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		pInetEx->Delete();
		if (pHttpFile != NULL)	delete pHttpFile;
		if (pHttpServer != NULL)	delete pHttpServer;
		if (pfile) fclose(pfile);
		INetSession.Close();
		return false;
	}

	if (bDownLoad)
	{
		//д.ini�ļ�

		if (!wdocx::MakeFile(szPath, index, szAttachmentFileName)) return false;

	}

	return true;
}


int  wdocx::IsNeedLoad(int index)
{
	int nMark = atoi(GetString("Mark", GetIniName(index)));
	int nInMark = atoi(GetString("Mark", GetIniName(index)));
	if (nMark + nInMark <= 0) DeleteDirFile(index);    //DeleteAll(index);

	//�ж������ļ��Ƿ��,���ļ���Ϊ��ʱ��˵���ļ��Ѿ���
	if (GetFileName("ini", "P_", index) == "") return -1;  //Ȩ��
	if (GetFileName("doc", "H_", index) == "") return -1;  //ͷ�ļ�
	if (GetFileName("doc", "T_", index) == "") return -1;  //ģ��
	if (GetFileName("doc", "D_", index) == "") return -1;  //�����ļ�
	if (GetFileName("bmp", "B_", index) == "") return -1;  //����

	//����������ļ�����ÿ�ζ�Ҫ����
	if (AfxGetApp()->GetProfileString("Telecom", "DeleteAllFile", "") != "") return true;

	if (GetString("IsNeedLoad", GetIniName(index)) == "1") return false;  //�����������Ҫ����

	return true;
}

BOOL wdocx::MakeFile(CString szFileName, int index, CString szAttachmentPath)
{

	FILE *pfile = NULL;
	pfile = fopen(szFileName, "rb");
	if (pfile == NULL)
	{
		MessageBox(NULL, "�������ص������ļ�ʧ�ܣ�������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}

	if (!SplitFile(pfile, GetFileName("ini", "P_", index), "HEADSTART", "HEADEND")) { fclose(pfile); return false; }

	if (!SplitFile(pfile, GetFileName("zip", "H_", index), "FILEHEADSTART", "FILEHEADEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "T_", index), "TMPSTART", "TMPEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "D_", index), "DATASTART", "DATAEND")) { fclose(pfile); return false; }
	if (!SplitFile(pfile, GetFileName("zip", "B_", index), "PICTURESTART", "PICTUREEND")) { fclose(pfile); return false; }
	if (pfile) fclose(pfile);

	if (index == 10)  //2003/11/26  �������������и����Ĺ���
	{  //�������и���
		if (!DeCompression(GetFileName("zip", "D_", index), szAttachmentPath, index)) return false;
	}
	else
	{

		//��ѹ���ļ�
		if (!DeCompression(GetFileName("zip", "H_", index), GetFileName("doc", "H_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "T_", index), GetFileName("doc", "T_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "D_", index), GetFileName("doc", "D_", index), index)) return false;
		if (!DeCompression(GetFileName("zip", "B_", index), GetFileName("bmp", "B_", index), index)) return false;

	}
	return true;
}
//����
BOOL wdocx::InsuerDocument(char * szHeader, char * szSomeString)
{
	int index = 2;
	CString szTextFile;
	CString szTemFile;
	CString szHeadFile;

	if (!wdocx::DocConnectionHttp(szHeader, strlen(szHeader), index)){ return false; }

	szTextFile = GetFileName("doc", "D_", index);
	if (szTextFile == "")  return false;
	szTemFile = GetFileName("doc", "T_", index);
	if (szTextFile == "") return false;
	szHeadFile = GetFileName("doc", "H_", index);
	if (szHeadFile == "") return false;

	if (!LastText(szTemFile, szHeadFile, szTextFile, szSomeString)) { DeleteFile(GetIniName(index)); return false; }

	WriteString("LastFileName", szTemFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	WriteString("Protect", "1", GetIniName(index));

	szFinalFile = GetFile("doc", "T_", index);


	return true;
}
BOOL wdocx::StampCover(char * szHeader)
{
	int index = 4;
	CString szTextFile;
	CString szPicture;

	if (!wdocx::DocConnectionHttp(szHeader, strlen(szHeader), index)){ return false; }  //�����ļ�
	szTextFile = GetFileName("doc", "D_", index);
	if (szTextFile == "") return false;
	szPicture = GetFileName("bmp", "B_", index);
	if (szPicture == "")  return false;




	if (!Stamp(szTextFile, szPicture)) { DeleteFile(GetIniName(index)); return false; }

	WriteString("LastFileName", szTextFile, GetIniName(index));
	WriteString("IsNeedLoad", "1", GetIniName(index));
	WriteString("Protect", "1", GetIniName(index));

	szFinalFile = GetFile("doc", "D_", index);

	return true;
}



BOOL wdocx::SendData(CString szHeader, CString szFileName, int index)
{
	if (!GetTheCabarcFile()) return false;

	CString szPath = szFileName;
	CString szCabFile;
	CString szCommand;

	int len = szFileName.Find("#|");
	if (len>0)
	{
		szPath = szFileName.Left(len);
		szFileName = szFileName.Mid(len + 2);
	}
	else//��������·����ȡ��
	{
		char buffer[256];
		memset(buffer, 0, sizeof(buffer));
		strcpy(buffer, szPath);
		szFileName = "";
		for (int i = strlen(buffer) - 1; i >= 0; i--)
		{
			if (buffer[i] == '\\') break;
			szFileName = buffer[i] + szFileName;
		}
	}
	szCommand.Format("%s\\unicom\\%s\\unicomOA", GetSysDirectory(), Dir[index]);


	CString szTemp;	szTemp = szFileName; szTemp.MakeUpper();
	int nrec = szTemp.Find(".ZIP");
	if (!OnFileCopy(szPath, szCommand)) return false;
	szCabFile.Format("%s\\unicom\\%s\\TempDoc.zip", GetSysDirectory(), Dir[index]);

	if (nrec<0)
	{ //ѹ��֮
		if (!Compression(szCabFile, szCommand)) return false;
		DeleteFile(szCommand);
	}
	else//��ѹ����ֻ����
	{
		szCabFile = szCommand;
	}

	FILE * pfile = NULL;
	pfile = fopen(szCabFile, "rb");

	if (pfile == NULL)
	{
		CString szInfo;
		szInfo.Format("�޷����ļ�%s�������ϴ�ʧ��,�����ԣ�", szFileName);
		MessageBox(NULL, szInfo, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
		return false;
	}
	DWORD nFileLen = 0;
	fseek(pfile, 0, SEEK_END);
	nFileLen = ftell(pfile);   //��ȡ�ļ�����
	rewind(pfile);           //ָ���Ƶ���ͷ
	if (nFileLen == 0) { if (pfile) fclose(pfile); DeleteFile(szCabFile); return true; }

	CString szInfo;//="f"+FileID+DBPath+"&^&%s#|#";

	szInfo.Format(szHeader, szFileName);
	int nlen = szInfo.GetLength();

	if (nFileLen <= 10 * 1000 * 1000)
	{
		char * buf = new char[nFileLen + nlen];
		memset(buf, 0, sizeof(buf));
		strcpy(buf, szInfo);
		fread((void*)(buf + nlen), 1, nFileLen, pfile);
		if (pfile) fclose(pfile);
		if (!wdocx::DocConnectionHttp(buf, nFileLen + nlen, index, 0)) //�ϴ��ļ�
		{
			delete buf;
			buf = NULL;
			return false;
		}
		delete buf;
	}
	else // 2003/7/9  �ϴ�����10M���ļ� 
	{
		DWORD FS = 5000000;
		int nindex = nFileLen / FS;
		bool bleave = 0;
		if (nFileLen%FS){ bleave = 1; nindex++; }

		char *buffer = (char*)malloc(FS + nlen + 100);

		AfxGetApp()->WriteProfileString("Telecom", "Large", "1");


		DWORD nAllCount = 0;
		nAllCount = nFileLen;
		for (int i = 1; i <= nindex; i++)
		{
			CString szSequence, szLast;
			if (i<10)		szSequence.Format("00%d", i);
			else if (i<100)	szSequence.Format("0%d", i);
			else			szSequence.Format("%d", i);
			szLast = "#" + szSequence;
			memset(buffer, 0, sizeof(buffer));
			strcpy(buffer, szInfo);


			if (i == nindex && bleave)
			{
				fread((void*)(buffer + nlen), 1, nAllCount, pfile);
				szLast += "y#";
				strcpy(buffer + nlen + nAllCount, szLast);
				nFileLen = nlen + nAllCount + szLast.GetLength();
			}
			else
			{
				fread((void*)(buffer + nlen), 1, FS, pfile);
				nAllCount -= FS;
				szLast += "n#";
				strcpy(buffer + nlen + FS, szLast);
				nFileLen = nlen + FS + szLast.GetLength();
			}

			//��������
			if (!wdocx::DocConnectionHttp(buffer, nFileLen, index, 0))
			{
				AfxGetApp()->WriteProfileString("Telecom", "Large", "0");
				if (pfile) fclose(pfile);
				free(buffer);
				return false;
			}
		}   //���ͽ���
		if (pfile) fclose(pfile);
		free(buffer);
		AfxGetApp()->WriteProfileString("Telecom", "Large", "0");
	}

	DeleteFile(szCabFile);

	return true;
}

BOOL wdocx::DownLoad(char * szInfo, char * szUpInfo, char * szFileName)
{
	AfxGetApp()->DoWaitCursor(1);

	int index = 9; //��ʾ��������
	CString szInformation;
	szInformation.Format(szInfo, szFileName);
	strcpy(szInfo, szInformation);
	CString szAttachFile;

	if (!wdocx::DocConnectionHttp(szInfo, strlen(szInfo), index)){ return false; }  //�����ļ�

	//���ɸ���
	char path[256];
	memset(path, 0, sizeof(path));
	strcpy(path, szFileName);
	CString szEx;
	for (int i = strlen(path) - 1; i>0; i--)
	{
		if (path[i] == '.') break;
		szEx = path[i] + szEx;
	}

	szAttachFile = GetFileName(szEx, "A_", index);
	if (szAttachFile == "") return false;

	if (!OnFileCopy(GetFileName("doc", "D_", index), szAttachFile))
	{
		MessageBox(NULL, "������������������", "ϵͳ��Ϣ", MB_OK | MB_ICONERROR);
		return false;
	}

	AfxGetApp()->DoWaitCursor(0);

	//�༭����
	if (!OpenAttachment(szAttachFile)) { DeleteFile(GetIniName(index)); return false; }

	CString szTempFileName;
	CString sztemp = szFileName;

	//	sztemp.Replace(" ","");  
	szTempFileName.Format("%s\\unicom\\%s\\%s", GetSysDirectory(), Dir[index], sztemp);
	DeleteFile(szTempFileName);
	if (!ReNameFile(szAttachFile, szTempFileName)) return false;
	szAttachFile = szTempFileName;
	//��������
	WriteString("IsNeedLoad", "1", GetIniName(index));     //����Щ��־λд�룬�Ա��ϴ�ʧ�ܺ��ٴδ�
	WriteString("LastFileName", szAttachFile, GetIniName(index));

	//����
	sztemp = szAttachFile + "#|" + sztemp;

	AfxGetApp()->DoWaitCursor(1);

	if (!wdocx::SendData(szUpInfo, sztemp, index)) return false;
	DeleteFile(szAttachFile);

	//DeleteAll(index);
	DeleteDirFile(index);
	AfxGetApp()->DoWaitCursor(0);

	return true;
}


int wdocx::DownLoadAllAttachmentEx(char * szInfo, CString szFileNames)
{
	int index = 10;
	char InfoBuf[256];
	//  memset(InfoBuf,0,sizeof(InfoBuf));

	//���ԭ������
	CString szDownLoadPath;
	szDownLoadPath.Format("%s\\unicom\\%s", GetSysDirectory(), Dir[index]);
	DeleteDataFile(szDownLoadPath);

	if (szFileNames == "")
	{
		MessageBox(NULL, "��ѡ��Ҫ���صĸ�����������!", "ϵͳ��Ϣ", MB_OK | MB_ICONWARNING);
		return false;
	}

	//	SetIpAndPort("172.16.10.21",81,"servlet/ULoadBDoc");
	CString szInformation;
	//ѡ������·��
	CBrowseDirDialog dlg;
	dlg.m_Title = "ѡ������·��";
	dlg.m_Path = "";
	if (dlg.DoBrowse() == 0) return 1;  //������

	CString szPath = dlg.m_Path;

	CStringArray szItem;
	CString szTempName;
	GetAllFileNames(szItem, szFileNames);
	int nCount = szItem.GetSize();  //��ȡҪ���ص��ļ���
	for (int i = 0; i<nCount; i++)
	{
		szTempName = szItem[i];

		if (!JudgeFileIgnoreOrNot(szPath, szTempName)) continue;
		memset(InfoBuf, 0, sizeof(InfoBuf));
		strcpy(InfoBuf, szInfo);
		szInformation.Format(InfoBuf, szItem[i]);
		memset(InfoBuf, 0, sizeof(InfoBuf));
		strcpy(InfoBuf, szInformation);
		szA_Name = szItem[i];   //���ļ����������Ա����غ����
		if (!wdocx::DocConnectionHttp(InfoBuf, strlen(InfoBuf), index, 1, szTempName)){ return false; }  //�����ļ�
	}
	return true;
}

BOOL wdocx::SendAttach(CString szInfo)
{

	int index = 9;
	static char BASED_CODE szFilter[] = "�����ļ�(*.*)|*.*|WPS�ļ�(*.WPS)|*.DOC|BMP�ļ�(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
	CString szfile1 = "", szfile2 = "";
	char BufFileNames[25600];
	memset(BufFileNames, 0, sizeof(BufFileNames));
	CFileDialog BrowseDialog(TRUE, "", "", OFN_ALLOWMULTISELECT, szFilter, NULL);

	BrowseDialog.m_ofn.lpstrFile = BufFileNames;         //2003/8/23 22:31
	BrowseDialog.m_ofn.nMaxFile = sizeof(BufFileNames);  //2003/8/23 22:31

	int nres = BrowseDialog.DoModal();

	if (nres == IDOK)
	{

		int ncount = 0;
		POSITION pos = BrowseDialog.GetStartPosition();
		CString file1 = szFileID;
		CString file2 = szFileID + "_dgc";


		AfxGetApp()->DoWaitCursor(1);

		for (;;)
		{
			CString FileName = BrowseDialog.GetNextPathName(pos);


			TRACE(FileName); TRACE("\n");
			if (FileName.Find(file1)>-1)
			{

				szfile1 = file1;
			}
			else if (FileName.Find(file2)>-1)
			{
				szfile2 = file2;

			}
			else //����
			{
				if (!wdocx::SendData(szInfo, FileName, index))   //2003/8/23 21:31
				{
					return false;
				}
			}
			if (pos == NULL) break;
			ncount++;
		}

	}

	if (szfile1 != "" && szfile2 != "")
	{
		MessageBox(NULL, szfile1 + szfile2 + "��ϵͳ�ļ���������������ٷ��ͣ������ļ��ѷ��ͳɹ���", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
	}
	return true;
}


BOOL wdocx::SendMailEx(CString szInfo, float fPart /*��KΪ��λ*/, float fTotal/*����Ϊ��λ*/)
{
	fTotal *= 1000;

	int index = 9;
	static char BASED_CODE szFilter[] = "�����ļ�(*.*)|*.*|WPS�ļ�(*.WPS)|*.DOC|BMP�ļ�(*.bmp)|*.bmp|GIF(*.gif)|*.gif||";
	CString szfile1 = "", szfile2 = "";
	char BufFileNames[25600];
	memset(BufFileNames, 0, sizeof(BufFileNames));
	CFileDialog BrowseDialog(TRUE, "", "", OFN_ALLOWMULTISELECT, szFilter, NULL);

	BrowseDialog.m_ofn.lpstrFile = BufFileNames;         //2003/8/23 22:31
	BrowseDialog.m_ofn.nMaxFile = sizeof(BufFileNames);  //2003/8/23 22:31
	CString file1 = szFileID;
	CString file2 = szFileID + "_dg";
	CStringArray  szItemNames;
	szItemNames.Add("test");
	szItemNames.RemoveAll();

	int nres = BrowseDialog.DoModal();
	if (nres == IDOK)
	{
		POSITION pos = BrowseDialog.GetStartPosition();

		AfxGetApp()->DoWaitCursor(1);
		DWORD  nAllSize = 0;
		for (;;)  //���淢�����ݵ�����
		{
			CString FileName = BrowseDialog.GetNextPathName(pos);
			DWORD nFileLen = GetFileLen(FileName);
			if (nFileLen<0) return false;   //���ļ���������
			nAllSize += nFileLen;
			szItemNames.Add(FileName);
			if (pos == NULL) break;
		}

		float fAllSize = (float)nAllSize / 1000;
		float fSize = (fTotal - fPart) / 1000;  //ת��ΪM

		if (fAllSize>(fTotal - fPart))
		{
			szItemNames.RemoveAll();
			CString szText;
			szText.Format("�ܵĸ�����СΪ%.2f�ף����Ѿ�������%.2f�ף������ٸ��ӳ���%.2f�׵ĸ�����", fTotal / 1000, fPart / 1000, fSize);
			MessageBox(NULL, szText, "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return false;
		}  //�жϽ���������������������

		//��������
		for (int i = 0; i<szItemNames.GetSize(); i++)
		{
			CString FileName = szItemNames.GetAt(i);
			if (FileName.Find(file1)>-1)
			{
				szfile1 = file1;
			}
			else if (FileName.Find(file2)>-1)
			{
				szfile2 = file2;
			}
			else //����
			{
				Sleep(100);
				if (!wdocx::SendData(szInfo, FileName, index))   //2003/8/23 21:31
				{
					return false;
				}
			}
		}
	}
	szItemNames.RemoveAll();

	if (szfile1 != "" || szfile2 != "")
	{
		MessageBox(NULL, szfile1 + szfile2 + "��ϵͳ�ļ���������������ٷ��ͣ������ļ��ѷ��ͳɹ���", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
	}
	return true;

}
