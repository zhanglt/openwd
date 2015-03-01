#ifndef OPENWD_WPS_H
#define OPENWD_WPS_H
/*----------------------------------------------
���ܣ�������в���Wps�ļ��Ĺ��ܺ���
ʱ�䣺2009��10��22
��д��zhanglt
------------------------------------------------*/
namespace wpsDoc {
#define EDIT   0    //�༭
#define MODIFY 1	//�޸�
#define READONLY   2	//���

	//���ĵ��ķ�ʽ
#define wpsOpenFormatAuto        0   //�Զ���ʽ�� 
#define wpsOpenFormatDocument    1   //�ĵ���ʽ�� 
#define wpsOpenFormatTemplate    2   //ģ�巽ʽ�� 
#define wpsOpenFormatRTF         3   //RTF ��ʽ�� 
#define wpsOpenFormatText        4   //�ı���ʽ�� 
#define wpsOpenFormatUnicodeText 5   //Unicode �ı���ʽ�� 
#define wpsOpenFormatWebPages    6   //��ҳ��ʽ�� 
	//�ĵ���������
#define wpsAllowOnlyComments     1   //�ĵ�������ʽΪ�޶�
#define wpsAllowOnlyFormFields   2
#define wpsAllowOnlyReading      3   //�ĵ������淽ʽΪֻ��
#define wpsAllowOnlyRevisions    0
#define wpsNoProtection          -1  //�ĵ����淽ʽΪ������




	//���ܣ���wps�ĵ�
	BOOL OpenWpsFile(char * szFileName/*wps�ļ���*/, char * szUserName, int nPower = 0/*Ȩ��*/, int bHaveTrace = 0);
	//���ܣ��򿪽����ĵ�
	BOOL OpenEmprstimoFile(char * szFileName, BOOL hide);

	//����
	BOOL Stamp(CString szFileName,/*��������ļ���*/ CString InserFileNames/*���й��µ��ļ���*/);
	//����
	BOOL LastText(CString szTempleteFileName,/*��������ļ���*/  CString szHeaderFileName/*�ļ�����*/, CString szDataFileName, CString szInfo);

	//����Ϊ���洦������
	//���
	//BOOL LookUpWps(CString szFileName,int bHaveTrace);
	//���Ĵ���
	BOOL EditFaxWps(CString szFileName, CString UserName, CString szHeader, int nPower, int bHaveTrace);
	//����  �����޸ģ�����ʾ�޸ĺۼ�
	BOOL FinalFaxWps(CString szFileName, CString  szHeader);
	//���Ķ��� �Ű�
	BOOL FinalFaxTextWps(CString szFileName, int nPower);
	//����
	BOOL StampFaxWps(CString szFileName, CString szStampFile);
	//���ļ���д����
	BOOL SetPortect(CString szFileName);


	//���ܣ�wps�ĵ�תΪHtml
	BOOL OpenHtmlFile(char * szFileName/*wps�ļ���*/, char * szUserName, int nPower = 0/*Ȩ��*/, int bHaveTrace = 0);
	//���ܣ��ӷ���������Wps�ļ�
	BOOL GetWpsFileFromServer(char* szInfo, char * szUsername = "", int bHaveTrace = 0);


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
	BOOL SendWpsFileToServer(char *szInfo = "", int index = 1);

	BOOL WpsConnectionHttp(char * TextBuf = "", DWORD nFileLen = 0, int index = 1, int bDownLoad = 1, CString szAttachmentFileName = "");
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

}
#endif