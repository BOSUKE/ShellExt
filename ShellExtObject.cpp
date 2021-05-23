// ShellExtObject.cpp : CShellExtObject の実装

#include "pch.h"
#include "ShellExtObject.h"
#pragma comment(lib, "mpr.lib")
#pragma comment(lib, "Userenv.lib")

namespace fs = std::filesystem;

// CShellExtObject
namespace {
  const wchar_t* SAVED_PATH_REG_KEY = L"SOFTWARE\\bosuke\\ShellExt";
  const wchar_t* SAVED_PATH_REG_NAME = L"SavedPath";

  void SavePath(const std::wstring& path)
  {
    RegSetKeyValueW(HKEY_CURRENT_USER, SAVED_PATH_REG_KEY, SAVED_PATH_REG_NAME, REG_SZ,
      path.c_str(), ((DWORD)path.size() + 1) * sizeof(wchar_t));
  }

  std::wstring LoadPath(void)
  {
    DWORD dataSize = 0;
    LONG retCode = RegGetValueW(HKEY_CURRENT_USER, SAVED_PATH_REG_KEY, SAVED_PATH_REG_NAME, RRF_RT_REG_SZ,
      NULL, NULL, &dataSize);
    if (retCode != ERROR_SUCCESS) return L"";

    std::wstring val;
    val.resize(dataSize / sizeof(wchar_t));
    retCode = RegGetValueW(HKEY_CURRENT_USER, SAVED_PATH_REG_KEY, SAVED_PATH_REG_NAME, RRF_RT_REG_SZ,
      NULL, &val[0], &dataSize);
    if (retCode != ERROR_SUCCESS) return L"";

    val.resize((dataSize / sizeof(wchar_t)) - 1);
    return val;
  }

  void SetTextToClipboard(const std::wstring& text)
  {
    size_t bufSize = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL h = GlobalAlloc(GHND, bufSize);
    if (h != NULL) {
      void* d = GlobalLock(h);
      if (d != NULL) {
        memcpy(d, text.c_str(), bufSize);
        GlobalUnlock(d);
      }
      if (OpenClipboard(NULL)) {
        EmptyClipboard();
        if (SetClipboardData(CF_UNICODETEXT, h) == NULL) {
          GlobalFree(h);
        }
        CloseClipboard();
      } else {
        GlobalFree(h);
      }
    }
  }

  std::wstring GetTextFromClipboard()
  {
    std::wstring ret;
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return ret;

    if (!OpenClipboard(NULL)) return ret;

    HGLOBAL h = GetClipboardData(CF_UNICODETEXT);
    if (h != NULL) {
      void* p = GlobalLock(h);
      if (p != NULL) {
        ret = (wchar_t*)p;
        GlobalUnlock(p);
      }
    }
    CloseClipboard();

    return ret;
  }

  std::wstring ConvertPathToUNC(const std::wstring& path)
  {
    if (path.length() == 0) return path;

    DWORD bufSize = MAX_PATH * sizeof(wchar_t);
    std::vector<uint8_t> buf(bufSize);

    DWORD ret = WNetGetUniversalNameW(path.c_str(), UNIVERSAL_NAME_INFO_LEVEL, &buf[0], &bufSize);
    if (ret == ERROR_MORE_DATA) {
      buf.resize(bufSize);
      ret = WNetGetUniversalNameW(path.c_str(), UNIVERSAL_NAME_INFO_LEVEL, &buf[0], &bufSize);
    }
    if (ret == NO_ERROR) {
      UNIVERSAL_NAME_INFOW* info = reinterpret_cast<UNIVERSAL_NAME_INFO*>(&buf[0]);
      return std::wstring(info->lpUniversalName);
    } else {
      return path;
    }
  }

  const std::wstring& GetUserProfilePath()
  {
    static std::wstring profilePath;

    if (!profilePath.empty()) {
      return profilePath;
    }

    HANDLE processHandle = GetCurrentProcess();
    HANDLE tokenHandle = NULL;
    BOOL openResult = OpenProcessToken(processHandle, TOKEN_QUERY, &tokenHandle);
    if (!openResult) return profilePath;

    wchar_t path[MAX_PATH] = { L'\0' };
    DWORD cch = MAX_PATH;
    if (GetUserProfileDirectoryW(tokenHandle, path, &cch)) {
      profilePath = path;
    }

    CloseHandle(tokenHandle);
    
    return profilePath;
  }

  fs::path GetPathFromClipboard()
  {
    std::wstring clipText = GetTextFromClipboard();
    if (clipText.empty()) return "";

    if (fs::exists(clipText)) {
      return clipText;
    }

    srell::wregex pattern(L"<(\\S+?)>(\r\n\\s*(\\S+))*");
    srell::wsmatch results;
    if (!srell::regex_search(clipText, results, pattern)) return "";

    if (results[3].matched) {
      fs::path path = fs::path(results[1]) / fs::path(results[3]);
      if (fs::exists(path)) return path;
      path = fs::path(GetUserProfilePath()) / path;
      if (fs::exists(path)) return path;
    }
    
    fs::path path = results[1];
    if (fs::exists(path)) return path;
    path = fs::path(GetUserProfilePath()) / path;
    if (fs::exists(path)) return path;
    
    return "";
  }

  void ExplorePath(const std::wstring& path, HWND hwnd)
  {
    if (path.empty()) return;

    if (fs::is_directory(path)) {
      ::ShellExecuteW(hwnd, L"explore", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
    } else {
      PIDLIST_ABSOLUTE pidl = ILCreateFromPathW(path.c_str());
      if (pidl != NULL) {
        SHOpenFolderAndSelectItems(pidl, 0, NULL, 0);
        ILFree(pidl);
      }
    }
  }

  bool CopyFile(const std::wstring& srcPath, const std::wstring& dstDir)
  {
    CComPtr<IFileOperation> fileOp;
    HRESULT hr = fileOp.CoCreateInstance(CLSID_FileOperation, NULL, CLSCTX_INPROC_SERVER);
    if (hr != S_OK) return false;

    CComPtr<IShellItem> src;
    hr = SHCreateItemFromParsingName(srcPath.c_str(), NULL, IID_PPV_ARGS(&src));
    if (hr != S_OK) return false;


    CComPtr<IShellItem> dst;
    hr = SHCreateItemFromParsingName(dstDir.c_str(), NULL, IID_PPV_ARGS(&dst));
    if (hr != S_OK) return false;

    hr = fileOp->CopyItem(src, dst, NULL, NULL);
    if (hr != S_OK) return false;

    hr = fileOp->PerformOperations();
    if (hr != S_OK) return false;

    BOOL aborted = FALSE;
    hr = fileOp->GetAnyOperationsAborted(&aborted);
    if ((hr != S_OK) || aborted) {
      return false;
    }
    return true;
  }

  void CopyFileToTempAndOpen(const std::wstring& filePath)
  {
    fs::path srcPath = filePath;
    fs::path dstDir = fs::temp_directory_path();
    fs::path dstPath = dstDir / srcPath.filename();
    if (!CopyFile(srcPath, dstDir)) return;
    ::ShellExecuteW(NULL, L"open", dstPath.c_str(), NULL, NULL, SW_SHOWNORMAL);
  }

}

void CShellExtObject::CopyPath(bool with_dir)
{

  if (mFilePathes.empty()) return;

  std::sort(mFilePathes.begin(), mFilePathes.end(),
    [](auto const& l, auto const& r) {
      bool l_is_Dir = fs::is_directory(l); 
      bool r_is_Dir = fs::is_directory(r);
      if (l_is_Dir == r_is_Dir) {
        return l < r;
      } else {
        return l_is_Dir;
      }
    });

  std::wstring copyText;
  const auto& endP = mFilePathes.end();
  for (auto p = mFilePathes.begin(); p != endP; ) {

    const fs::path firstPath(*p);
    const auto orgRoot = firstPath.root_path();
    const auto orgDir = firstPath.parent_path();

    std::wstring orgDirStr = orgDir.string<wchar_t>();
    if (orgDirStr.empty() || (orgDirStr[orgDirStr.size() - 1] != L'\\')) {
      orgDirStr += L"\\";
    }

    std::wstring newDirStr = orgDirStr;
    const std::wstring& profilePath = GetUserProfilePath();
    if ((profilePath.size() != newDirStr.size())
      && (_wcsnicmp(profilePath.c_str(), newDirStr.c_str(), profilePath.size()) == 0)) {
      newDirStr = newDirStr.substr(profilePath.size() + 1);
    } else {
      newDirStr = ConvertPathToUNC(newDirStr);
    }
    if (newDirStr.empty() || (newDirStr[newDirStr.size() - 1] != L'\\')) {
      newDirStr += L"\\";
    }

    size_t subOffset = orgDirStr.size();
    if (with_dir) {
      for ( ; p != endP; ++p) {
        if (orgDir != fs::path(*p).parent_path()) {
          break;
        }
        copyText += L"<" + newDirStr + p->substr(subOffset);
        if (fs::is_directory(*p)) {
          copyText += L"\\";
        }
        copyText += L">\r\n";
      }
    } else {
      copyText += L"<" + newDirStr + L">\r\n";
      for (; p != endP; ++p) {
        if (orgDir != fs::path(*p).parent_path()) {
          break;
        }
        copyText += p->substr(subOffset);
        if (fs::is_directory(*p)) {
          copyText += L"\\";
        }
        copyText += L"\r\n";
      }
    }
  }
    ::SetTextToClipboard(copyText);
}

void CShellExtObject::SavePath()
{
  if (mFilePathes.size() != 1) return;
  ::SavePath(mFilePathes[0]);
}

void CShellExtObject::CreateShortcut()
{
  const std::wstring savedPath = LoadPath();
  if (savedPath.empty()) return;
  if (mFilePathes.size() != 1) return;
  if (!fs::is_directory(mFilePathes[0])) return;

  const fs::path dstPath = ConvertPathToUNC(savedPath);
  fs::path dstLastNode;
  for (const auto& n : dstPath) {
    dstLastNode = n;
  }
  const fs::path srcDir = mFilePathes[0];
  const fs::path srcPath = (srcDir / dstLastNode).replace_extension("lnk");
  
  CComPtr<IShellLink> shellLink;
  HRESULT hr = shellLink.CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER);
  if (hr != S_OK) return;

  shellLink->SetPath(dstPath.c_str());
  if (!fs::is_directory(dstPath)) {
    shellLink->SetWorkingDirectory(dstPath.parent_path().c_str());
  }

  CComPtr<IPersistFile> iFile;
  iFile = shellLink;
  if (iFile == NULL) return;

  iFile->Save(srcPath.c_str(), TRUE);
}

void CShellExtObject::ExploreSavedPath(HWND hwnd)
{
  const std::wstring savedPath = LoadPath();
  ExplorePath(savedPath, hwnd);
}

void CShellExtObject::OpenClipBoardPath(HWND hwnd)
{
  const auto path = GetPathFromClipboard();
  if (path.empty()) return;
  ::ShellExecuteW(hwnd, L"open", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
}

void CShellExtObject::ExploreClipBoardPath(HWND hwnd)
{
  const auto path = GetPathFromClipboard();
  ExplorePath(path.string<wchar_t>(), hwnd);
}

STDMETHODIMP CShellExtObject::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID)
{

  mFilePathes.clear();
  mClickFolder = false;

  if (pdtobj != NULL) {

    STGMEDIUM medium;
    FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    if (FAILED(pdtobj->GetData(&fmt, &medium))) {
      return E_INVALIDARG;
    }

    HDROP hdrop = (HDROP)::GlobalLock(medium.hGlobal);
    if (hdrop == NULL) {
      return E_INVALIDARG;
    }

    UINT count = DragQueryFileW(hdrop, 0xFFFFFFFF, NULL, 0);
    if (count != 0) {
      mFilePathes.reserve(count);
      for (UINT i = 0; i < count; i++) {

        UINT len = DragQueryFileW(hdrop, i, NULL, 0);
        if (len == 0) break;

        std::wstring fpath;
        fpath.resize((size_t)len + 1);
        if (DragQueryFileW(hdrop, i, fpath.data(), len + 1)) {
          fpath.resize(len);
          mFilePathes.push_back(fpath);
        }
      }
    }
    GlobalUnlock(hdrop);
    ReleaseStgMedium(&medium);
  } else {
    wchar_t fpath[MAX_PATH];
    if (!SHGetPathFromIDListW(pidlFolder, fpath)) {
      return E_INVALIDARG;
    }
    if (fs::is_directory(fpath)) {
      mFilePathes.push_back(fpath);
      mClickFolder = true;
    }
  } 
  return S_OK;
}

STDMETHODIMP CShellExtObject::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax)
{
  return E_INVALIDARG;
}

STDMETHODIMP CShellExtObject::InvokeCommand(CMINVOKECOMMANDINFO* pici)
{
  if (HIWORD(pici->lpVerb) != 0) {
    return E_INVALIDARG;
  }
  int menuItemIndex = LOWORD(pici->lpVerb);
  if (menuItemIndex >= MAX_MENU_ITEM) {
    return E_INVALIDARG;
  }
  int menuItem = mMenuItem[menuItemIndex];
  switch (menuItem) {
  case COPY_PATH_NO_DIR:
    CopyPath(false);
    break;
  case COPY_PATH_WITH_DIR:
    CopyPath(true);
    break;
  case SAVE_PATH:
    SavePath();
    break;
  case CREATE_SHORTCUT:
    CreateShortcut();
    break;
  case EXPLORE_SAVED_PATH:
    ExploreSavedPath(pici->hwnd);
    break;
  case OPEN_CLIPBOARD_PATH:
    OpenClipBoardPath(pici->hwnd);
    break;
  case EXPLORE_CLIPBOARD_PATH:
    ExploreClipBoardPath(pici->hwnd);
    break;
  case COPY_AND_OPEN:
    CopyAndOpen();
    break;
  default:
    return E_INVALIDARG;
  }
  return S_OK;
}

STDMETHODIMP CShellExtObject::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
  static int hoge = 0;
  if (uFlags & CMF_DEFAULTONLY) {
    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
  }

  HMENU submenu = CreatePopupMenu();
  UINT id = idCmdFirst;
  UINT pos = 0;

  auto addItem = [&](int menuItem, const wchar_t* str) {
    InsertMenuW(submenu, pos, MF_BYPOSITION, id, str);
    mMenuItem[pos] = menuItem;
    pos++;
    id++;
  };
  addItem(COPY_PATH_NO_DIR, L"パスコピー(ディレクトリ別)");
  addItem(COPY_PATH_WITH_DIR, L"パスコピー(ディレクトリ付)");
  if (mFilePathes.size() == 1) {
    addItem(SAVE_PATH, L"パスを覚える");
  }

  if (mClickFolder) {
    const std::wstring savedPath = LoadPath();
    if (!savedPath.empty()) {
      addItem(CREATE_SHORTCUT, L"覚えた奴のショートカット作成");
      addItem(EXPLORE_SAVED_PATH, L"覚えた奴の場所を開く");
    }
    if (!GetPathFromClipboard().empty()) {
      addItem(OPEN_CLIPBOARD_PATH, L"ClipBoardの奴を開く");
      addItem(EXPLORE_CLIPBOARD_PATH, L"ClipBoardの奴の場所を開く");
    }
  } else {
    if (mFilePathes.size() == 1) {
      fs::path p = mFilePathes[0];
      if (fs::exists(p) && fs::is_regular_file(p)) {
        addItem(COPY_AND_OPEN, L"コピーして開く");
      }
    }
  }
  std::fill_n(mMenuItem + pos, MAX_MENU_ITEM - pos, MAX_MENU_ITEM);

  MENUITEMINFO mii;
  mii.cbSize = sizeof(mii);
  mii.fMask = MIIM_SUBMENU | MIIM_STRING | MIIM_ID;
  mii.wID = id++;
  mii.hSubMenu = submenu;
  mii.dwTypeData = L"ShellExt";
  InsertMenuItemW(hmenu, indexMenu, TRUE, &mii);

  return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, id - idCmdFirst);
}

void CShellExtObject::CopyAndOpen()
{
  if (mFilePathes.size() != 1) return;
  std::thread t(CopyFileToTempAndOpen, mFilePathes[0]);
  t.detach();
}
