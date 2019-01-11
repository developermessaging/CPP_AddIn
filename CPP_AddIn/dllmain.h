// dllmain.h : Declaration of module class.

class CCPP_AddInModule : public ATL::CAtlDllModuleT< CCPP_AddInModule >
{
public :
	DECLARE_LIBID(LIBID_CPP_AddInLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CPP_AddIn, "{f3eda8a5-30af-4dc3-875b-075167f21b10}")
};

extern class CCPP_AddInModule _AtlModule;
