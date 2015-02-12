// dllmain.h : 模块类的声明。

class COpenWDModule : public ATL::CAtlDllModuleT< COpenWDModule >
{
public :
	DECLARE_LIBID(LIBID_OpenWDLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_OPENWD, "{716613F8-CB81-4989-AA91-DB1BAC8F8F3B}")
};

extern class COpenWDModule _AtlModule;
