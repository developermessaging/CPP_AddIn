// dllmain.h : Declaration of module class.

class CCPPComAddInModule : public ATL::CAtlDllModuleT< CCPPComAddInModule >
{
public :
	DECLARE_LIBID(LIBID_CPPComAddInLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CPPCOMADDIN, "{18CEFF5F-832A-49AC-8E53-579A51D524A5}")
};

extern class CCPPComAddInModule _AtlModule;
