class CComUtilityModule : public ATL::CAtlDllModuleT< CComUtilityModule >
{
public :
    DECLARE_LIBID(LIBID_ComUtilityLib)
    DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMUTILITY, "{154CCAE9-4F4D-47CC-B8CB-D26C760739F1}")
};

extern class CComUtilityModule _AtlModule;
