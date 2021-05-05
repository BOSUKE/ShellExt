// dllmain.h : モジュール クラスの宣言です。

class CShellExtModule : public ATL::CAtlDllModuleT< CShellExtModule >
{
public :
	DECLARE_LIBID(LIBID_ShellExtLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SHELLEXT, "{48ad19a7-9fa2-4f57-9b77-8d42feea8bf3}")
};

extern class CShellExtModule _AtlModule;
