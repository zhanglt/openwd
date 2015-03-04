#pragma once

#define  __BUFFER_SIZE 1024

class CHttpFile;

class CHttpFileClient
{
public:
	CHttpFileClient(void);
	~CHttpFileClient(void);

public:
	BOOL UploadFile(LPCTSTR szRemoteURI, LPCTSTR szLocalPath);
	BOOL DownLoadFile(LPCTSTR szRemoteURI, LPCTSTR szLocalPath);
	BOOL DeleteFile(LPCTSTR szRemoteURI);
	BOOL CanWebsiteVisit(CString sURI);

private:
	BOOL UseHttpSendReqEx(CHttpFile* httpFile, LPCTSTR szLocalFile);
};