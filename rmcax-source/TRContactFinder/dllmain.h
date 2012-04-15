// dllmain.h : Declaration of module class.

class CTRContactFinderModule : public CAtlDllModuleT< CTRContactFinderModule >
{
public :
	DECLARE_LIBID(LIBID_TRContactFinderLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TRCONTACTFINDER, "{DBBEECDF-A5AA-43A8-BA3A-261EBB1A83F2}")
};

extern class CTRContactFinderModule _AtlModule;
