
/*----------------------------------------------
���ܣ�word�ļ����ܺ���
------------------------------------------------*/
namespace wdocx {
#define EDIT        0   //�༭״̬
#define MODIFY      1   //�޸�״̬
#define READONLY    2	//���״̬
#define FINALEDIT   3	//�����༭���Զ����ܺۼ�

	//���ܣ���word�ĵ�
	BOOL OpenWordFile(char * szFileName/*word�ļ���*/, char * szUserName, int nPower = 0/*Ȩ��*/, int bHaveTrace = 0/*�ۼ�*/);
	//���ܣ��򿪽����ĵ�
	//BOOL OpenWordEmprstimoFile (char * szFileName,BOOL hide);
	//����
	BOOL Stamp(CString szFileName,/*��������ļ���*/ CString InserFileNames/*���й��µ��ļ���*/);
	//����
	BOOL LastText(CString szTempleteFileName,/*��������ļ���*/  CString szHeaderFileName/*�ļ�����*/, CString szDataFileName, CString szInfo);
	//����Ϊ���洦������
	//���
	//BOOL LookUpWord(CString szFileName,int bHaveTrace);
	//���Ĵ���
	BOOL EditFaxWord(CString szFileName, CString UserName, CString szHeader, int nPower, int bHaveTrace);
	//����  �����޸ģ�����ʾ�޸ĺۼ�
	BOOL FinalFaxWord(CString szFileName, CString  szHeader);
	//���Ķ��� �Ű�
	BOOL FinalFaxTextWord(CString szFileName, int nPower);
	//����
	BOOL StampFaxWord(CString szFileName, CString szStampFile);
	//���ļ���д����
	BOOL SetPortect(CString szFileName);
	//���ܣ�word�ĵ�תΪHtml
	BOOL OpenHtmlFile(char * szFileName/*word�ļ���*/, char * szUserName, int nPower = 0/*Ȩ��*/, int bHaveTrace = 0);
	//���ܣ��ӷ���������Doc�ļ�
	BOOL GetDocFileFromServer(char* szInfo, char * szUsername = "", int bHaveTrace = 0);
	//����
	BOOL StampFaxEx(char * szInfo);
	BOOL FinalTextEx(char *szInfo, int nPower);
	//���Ĵ���
	BOOL EditFaxEx(char *szInfo, char * szHeader, char * UserName, int nPower, int bHaveTrace);
	//����
	BOOL FinalFaxEx(char * szInfo, char *szHeader);
	//���Ķ���
	BOOL FinalFaxTextEx(char * szInfo, int nPower);
	//���ܣ���������ϴ��ļ�
	BOOL SendDocFileToServer(char *szInfo = "", int index = 1);
	int  IsNeedLoad(int index);
	BOOL MakeFile(CString szFileName, int index, CString szAttachmentPath);
	//���ܣ���������ϴ��ļ�
	BOOL InsuerDocument(char * szHeader, char * szSomeString = "");
	BOOL StampCover(char * szHeader);
	BOOL SendData(CString szHeader, CString szFileName, int index);
	BOOL DownLoad(char * szInfo, char * szUpInfo, char * szFileName);
	int DownLoadAllAttachmentEx(char * szInfo, CString szFileNames);
	BOOL SendAttach(CString szInfo);
	BOOL SendMailEx(CString szInfo, float fPart /*��KΪ��λ*/, float fTotal/*����Ϊ��λ*/);
	BOOL DocConnectionHttp(char * TextBuf = "", DWORD nFileLen = 0, int index = 1, int bDownLoad = 1, CString szAttachmentFileName = "");
}