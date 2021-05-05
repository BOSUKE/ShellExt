// ShellExtObject.h : CShellExtObject の宣言

#pragma once
#include "resource.h"       // メイン シンボル



#include "ShellExt_i.h"
#include <shlobj.h>
#include <comdef.h>



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "DCOM の完全サポートを含んでいない Windows Mobile プラットフォームのような Windows CE プラットフォームでは、単一スレッド COM オブジェクトは正しくサポートされていません。ATL が単一スレッド COM オブジェクトの作成をサポートすること、およびその単一スレッド COM オブジェクトの実装の使用を許可することを強制するには、_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA を定義してください。ご使用の rgs ファイルのスレッド モデルは 'Free' に設定されており、DCOM Windows CE 以外のプラットフォームでサポートされる唯一のスレッド モデルと設定されていました。"
#endif

using namespace ATL;


// CShellExtObject

class ATL_NO_VTABLE CShellExtObject :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CShellExtObject, &CLSID_ShellExtObject>,
	public IShellExtInit,
	public IContextMenu
{
public:
	CShellExtObject()
	:mClickFolder(false){
	}

DECLARE_REGISTRY_RESOURCEID(106)

DECLARE_NOT_AGGREGATABLE(CShellExtObject)

BEGIN_COM_MAP(CShellExtObject)
	COM_INTERFACE_ENTRY(IShellExtInit)
	COM_INTERFACE_ENTRY(IContextMenu)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

protected:
	bool mClickFolder;
	std::vector<std::wstring> mFilePathes;

	void CopyPath(bool with_dir);
	void SavePath();
	void CreateShortcut();
	void ExploreSavedPath(HWND hwnd);
	void OpenClipBoardPath(HWND hwnd);
	void ExploreClipBoardPath(HWND hwnd);

public:
	// IShellExtInit
	STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID);
	// IContextMenu
	STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax);
	STDMETHODIMP InvokeCommand(CMINVOKECOMMANDINFO* pici);
	STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT  indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
};

OBJECT_ENTRY_AUTO(__uuidof(ShellExtObject), CShellExtObject)
