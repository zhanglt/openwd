#if !defined(AFX_CONNECTIONADVISOR_H__10B0E294_B541_11D2_A0A2_0080C7F3B56B__INCLUDED_)
#define AFX_CONNECTIONADVISOR_H__10B0E294_B541_11D2_A0A2_0080C7F3B56B__INCLUDED_

/*----------------------------------------------------------------------------*/

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

/*----------------------------------------------------------------------------*/

class CConnectionAdvisor  
{
public:
	CConnectionAdvisor(REFIID iid);
	BOOL Advise(IUnknown* pSink, IUnknown* pSource);
	BOOL Unadvise();
	virtual ~CConnectionAdvisor();

private:
	CConnectionAdvisor();
	CConnectionAdvisor(const CConnectionAdvisor& ConnectionAdvisor);
	REFIID m_iid;
	IConnectionPoint* m_pConnectionPoint;
	DWORD m_AdviseCookie;
};

/*----------------------------------------------------------------------------*/

#endif // !defined(AFX_CONNECTIONADVISOR_H__10B0E294_B541_11D2_A0A2_0080C7F3B56B__INCLUDED_)
