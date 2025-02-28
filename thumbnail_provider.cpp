#include <atlbase.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <shlobj_core.h>
#include <shlwapi.h>
#include <stdexcept>
#include <strsafe.h>
#include <thumbcache.h>
#include <vector>
#include <wincodec.h>
#include <windows.h>

#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "user32.lib")

// Change it to your own extension.
#define FILE_EXTENSION "plasticity"
#define FILE_EXTENSIONW L".plasticity"

// Generate new UUID with `uuidgen -c` command in the terminal.
#define THUMBNAIL_HANDLER_GUID L"{6C68F6DE-04B0-4524-8276-106DB61B06B7}"

#define CLSID_THUMBNAIL_HANDLER_GUID L"{E357FCCD-A995-4576-B01F-234630154E96}"

const uint32_t JSON_MAGIC = 0x4e4f534a;
const uint32_t BIN_MAGIC = 0x004e4942;
const uint32_t THMB_MAGIC = 0x424d4854;

HMODULE g_hModule;

class GdiPlusScope {
public:
  ULONG_PTR token;

  GdiPlusScope() : token(0) {
    Gdiplus::GdiplusStartupInput input;
    status = Gdiplus::GdiplusStartup(&token, &input, NULL);
    if (status != Gdiplus::Ok) {
      throw std::runtime_error("GDI+ initialization failed");
    }
  }

  ~GdiPlusScope() {
    if (status == Gdiplus::Ok) {
      Gdiplus::GdiplusShutdown(token);
    }
  }

private:
  Gdiplus::Status status;
};

class __declspec(uuid("6C68F6DE-04B0-4524-8276-106DB61B06B7")) ThumbnailProvider :
  public IInitializeWithFile,
  public IInitializeWithStream,
  public IThumbnailProvider,
  public IExtractIconW {
public:
  IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
    static const QITAB qit[] = {
      QITABENT(ThumbnailProvider, IInitializeWithFile),
      QITABENT(ThumbnailProvider, IInitializeWithStream),
      QITABENT(ThumbnailProvider, IThumbnailProvider),
      QITABENT(ThumbnailProvider, IExtractIconW),
      { 0 },
    };
    return QISearch(this, qit, riid, ppv);
  }

  IFACEMETHODIMP_(ULONG) AddRef() {
    ULONG cRef = InterlockedIncrement(&_cRef);
    return  cRef;
  }

  IFACEMETHODIMP_(ULONG) Release() {
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (!cRef) {
      delete this;
    }
    return cRef;
  }

  // IInitializeWithStream
  IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode) {
    _pStream = pStream;

    return S_OK;
  }

  // IInitializeWithFile
  IFACEMETHODIMP Initialize(LPCWSTR pszFilePath, DWORD grfMode) {
    HRESULT hr = E_UNEXPECTED;
    hr = StringCchCopyW(_szFilePath, ARRAYSIZE(_szFilePath), pszFilePath);

    if (SUCCEEDED(hr)) {
      _pszFilePath = _szFilePath;
    }

    return hr;
  }

  std::vector<BYTE> ParseThumbNailFromFileData(const std::vector<BYTE>& fileData) {
    size_t offset = 0;

    std::string magic(fileData.begin(), fileData.begin() + 10);
    if (magic != FILE_EXTENSION) {
      return std::vector<BYTE>();
    }
    offset += 10;

    uint32_t version = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
    if (version != 1) {
      return std::vector<BYTE>();
    }
    offset += 4;

    uint32_t length = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
    if (fileData.size() < length) {
      return std::vector<BYTE>();
    }
    offset += 4;

    {
      uint32_t jsonLength = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      offset += 4;
      uint32_t magic = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      if (magic != JSON_MAGIC) {
        return std::vector<BYTE>();
      }
      offset += 4;
      offset += jsonLength;
    }

    {
      uint32_t binLength = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      offset += 4;
      uint32_t magic = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      if (magic != BIN_MAGIC) {
        return std::vector<BYTE>();
      }
      offset += 4;
      offset += binLength;
    }

    if (offset < fileData.size()) {
      uint32_t thumbLength = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      offset += 4;
      uint32_t magic = *reinterpret_cast<const uint32_t*>(&fileData[offset]);
      if (magic != THMB_MAGIC) return std::vector<BYTE>();
      offset += 4;
      return std::vector<BYTE>(fileData.begin() + offset, fileData.begin() + offset + thumbLength);
    }

    return std::vector<BYTE>();
  }

  HBITMAP PNGDataToHBITMAP(const std::vector<BYTE>& fileData) {
      std::vector<BYTE> pngData = ParseThumbNailFromFileData(fileData);
      if (pngData.empty()) {
        return NULL;
      }

      IStream* stream = SHCreateMemStream(pngData.data(), pngData.size());
      if (!stream) {
        return NULL;
      }

      Gdiplus::Bitmap* bitmap = Gdiplus::Bitmap::FromStream(stream);
      stream->Release();
      if (!bitmap) {
        return NULL;
      }

      HBITMAP hBitmap;
      bitmap->GetHBITMAP(Gdiplus::Color::Transparent, &hBitmap);
      delete bitmap;

      return hBitmap;
  }

  // IThumbnailProvider
  IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha) {
    try {
      ULONG_PTR gdiplusToken;
      Gdiplus::GdiplusStartupInput gdiplusStartupInput;
      Gdiplus::Status status = Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

      if (status != Gdiplus::Ok) {
        return E_FAIL;
      }

      *phbmp = NULL;
      std::vector<BYTE> fileData;
      if (_pStream) {
        // Read from stream
        STATSTG stat;
        if (SUCCEEDED(_pStream->Stat(&stat, STATFLAG_NONAME))) {
          fileData.resize(stat.cbSize.QuadPart);
          ULONG bytesRead;
          _pStream->Read(fileData.data(), stat.cbSize.QuadPart, &bytesRead);

          HBITMAP screenshot = PNGDataToHBITMAP(fileData);
          *phbmp = screenshot;
        }
      } else if (_pszFilePath) {
        // Read from file path
        HANDLE hFile = CreateFileW(
          _pszFilePath,
          GENERIC_READ,
          FILE_SHARE_READ,
          NULL,
          OPEN_EXISTING,
          FILE_ATTRIBUTE_NORMAL,
          NULL);

        if (hFile != INVALID_HANDLE_VALUE) {
          DWORD fileSize = GetFileSize(hFile, NULL);
          fileData.resize(fileSize);
          DWORD bytesRead;
          ReadFile(
            hFile,
            fileData.data(),
            fileSize,
            &bytesRead,
            NULL);

          CloseHandle(hFile);

          HBITMAP screenshot = PNGDataToHBITMAP(fileData);
          *phbmp = screenshot;
        }
      }

      Gdiplus::GdiplusShutdown(gdiplusToken);
    } catch(...) {
      return E_FAIL;
    }

    return *phbmp ? S_OK : E_FAIL;
  }

  // IExtractIcon implementation
  IFACEMETHODIMP GetIconLocation(UINT uFlags,
                                 LPWSTR  pszIconFile,
                                 UINT cchMax,
                                 int* piIndex,
                                 UINT* pwFlags) override {
    // Tell Windows we'll extract icons ourselves
    *pwFlags = GIL_NOTFILENAME | GIL_DONTCACHE;
    *piIndex = 0;
    return S_OK;
  }

  IFACEMETHODIMP Extract(LPCWSTR pszFile,
                         UINT nIconIndex,
                         HICON* phiconLarge,
                         HICON* phiconSmall,
                         UINT nIconSize) {
    std::wstring wPath(pszFile);
    std::string path2(wPath.begin(), wPath.end());

    try {
      std::vector<BYTE> fileData;

      // Read file data
      if (_pStream) {
        // Read from stream
        STATSTG stat;
        if (SUCCEEDED(_pStream->Stat(&stat, STATFLAG_NONAME))) {
          fileData.resize(stat.cbSize.QuadPart);
          ULONG bytesRead;
          _pStream->Read(fileData.data(), stat.cbSize.QuadPart, &bytesRead);
        }
      } else if (_pszFilePath) {
        // Read from file path
        HANDLE hFile = CreateFileW(
          _pszFilePath,
          GENERIC_READ,
          FILE_SHARE_READ,
          NULL,
          OPEN_EXISTING,
          FILE_ATTRIBUTE_NORMAL,
          NULL);

        if (hFile != INVALID_HANDLE_VALUE) {
          DWORD fileSize = GetFileSize(hFile, NULL);
          fileData.resize(fileSize);
          DWORD bytesRead;
          ReadFile(
            hFile,
            fileData.data(),
            fileSize,
            &bytesRead,
            NULL);
          CloseHandle(hFile);
        }
      }

      if (fileData.empty()) {
        return E_FAIL;
      }

      GdiPlusScope gdiPlus;
      // Get PNG data
      std::vector<BYTE> pngData = ParseThumbNailFromFileData(fileData);
      if (pngData.empty()) {
        Gdiplus::GdiplusShutdown(gdiPlus.token);
        return E_FAIL;
      }

      // Create stream from PNG data
      IStream* stream = SHCreateMemStream(pngData.data(), pngData.size());
      if (!stream) {
        Gdiplus::GdiplusShutdown(gdiPlus.token);
        return E_FAIL;
      }

      // Load bitmap
      Gdiplus::Bitmap* originalBitmap = Gdiplus::Bitmap::FromStream(stream);
      stream->Release();

      if (!originalBitmap) {
        Gdiplus::GdiplusShutdown(gdiPlus.token);
        return E_FAIL;
      }

      // Create large icon (32x32)
      Gdiplus::Bitmap* largeBitmap = new Gdiplus::Bitmap(32, 32, PixelFormat32bppARGB);
      Gdiplus::Graphics graphics(largeBitmap);
      graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
      graphics.DrawImage(originalBitmap, 0, 0, 32, 32);
      largeBitmap->GetHICON(phiconLarge);
      delete largeBitmap;

      // Create small icon (16x16)
      Gdiplus::Bitmap* smallBitmap = new Gdiplus::Bitmap(16, 16, PixelFormat32bppARGB);
      Gdiplus::Graphics graphics1(smallBitmap);
      graphics1.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
      graphics1.DrawImage(originalBitmap, 0, 0, 16, 16);
      smallBitmap->GetHICON(phiconSmall);
      delete smallBitmap;

      delete originalBitmap;

      return S_OK;
    }
    catch (...) {
      return E_FAIL;
    }
  }

  ThumbnailProvider() : _cRef(1), _pszFilePath(NULL) {}

private:
  ~ThumbnailProvider() {}
  long _cRef;
  WCHAR _szFilePath[MAX_PATH];
  LPCWSTR _pszFilePath;
  CComPtr<IStream> _pStream;
};

class ThumbnailProviderFactory : public IClassFactory {
public:
  static HRESULT CreateInstance(REFIID riid, void **ppv) {
    *ppv = NULL;
    ThumbnailProviderFactory *pFactory = new (std::nothrow) ThumbnailProviderFactory();
    HRESULT hr = pFactory ? S_OK : E_OUTOFMEMORY;

    if (SUCCEEDED(hr)) {
      hr = pFactory->QueryInterface(riid, ppv);
      pFactory->Release();
    }

    return hr;
  }

private:
  ThumbnailProviderFactory() : _cRef(1) {}
  virtual ~ThumbnailProviderFactory() {}

  IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv) {
    static const QITAB qit[] = {
      QITABENT(ThumbnailProviderFactory, IClassFactory),
      { 0 },
    };

    return QISearch(this, qit, riid, ppv);
  }

  IFACEMETHODIMP_(ULONG) AddRef() {
    return InterlockedIncrement(&_cRef);
  }

  IFACEMETHODIMP_(ULONG) Release() {
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (!cRef) {
      delete this;
    }

    return cRef;
  }

  // IClassFactory
  IFACEMETHODIMP CreateInstance(IUnknown *punkOuter, REFIID riid, void **ppv) {
    *ppv = NULL;
    HRESULT hr = CLASS_E_NOAGGREGATION;

    if (punkOuter == NULL) {
      hr = E_OUTOFMEMORY;
      ThumbnailProvider *pProvider = new (std::nothrow) ThumbnailProvider();

      if (pProvider) {
        hr = pProvider->QueryInterface(riid, ppv);
        pProvider->Release();
      }
    }

    return hr;
  }

  IFACEMETHODIMP LockServer(BOOL fLock) {
    if (fLock) {
      InterlockedIncrement(&_cRef);
    } else {
      InterlockedDecrement(&_cRef);
    }

    return S_OK;
  }

private:
  long _cRef;
};

// DLL exports
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
  g_hModule = hModule;

  return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv) {
  HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

  if (IsEqualCLSID(rclsid, __uuidof(ThumbnailProvider))) {
    hr = ThumbnailProviderFactory::CreateInstance(riid, ppv);
  }

  return hr;
}

STDAPI DllCanUnloadNow() {
  return S_OK;
}

STDAPI DllRegisterServer(void) {
  WCHAR szModule[MAX_PATH];
  GetModuleFileNameW(g_hModule, szModule, ARRAYSIZE(szModule));

  // Main CLSID registration
  HKEY hKeyLM;
  HKEY hKey;
  if (RegCreateKeyExW(
    HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\CLSID\\" THUMBNAIL_HANDLER_GUID,
    0,
    NULL,
    REG_OPTION_NON_VOLATILE,
    KEY_WRITE,
    NULL,
    &hKeyLM,
    NULL) == ERROR_SUCCESS
  ) {
    // InProcServer32
    HKEY hKeyServer;
    if (RegCreateKeyExW(
      hKeyLM,
      L"InProcServer32",
      0,
      NULL,
      REG_OPTION_NON_VOLATILE,
      KEY_WRITE,
      NULL,
      &hKeyServer,
      NULL) == ERROR_SUCCESS
    ) {
      RegSetValueExW(
        hKeyServer,
        NULL,
        0,
        REG_SZ,
        (BYTE*)szModule,
        (wcslen(szModule) + 1) * sizeof(WCHAR));

      RegSetValueExW(
        hKeyServer,
        L"ThreadingModel",
        0,
        REG_SZ,
        (BYTE*)L"Both",
        sizeof(L"Both"));
      RegCloseKey(hKeyServer);
    }

    // Implemented Categories
    RegCreateKeyExW(
      hKeyLM,
      L"Implemented Categories\\" CLSID_THUMBNAIL_HANDLER_GUID,
      0,
      NULL,
      REG_OPTION_NON_VOLATILE,
      KEY_WRITE,
      NULL,
      &hKey,
      NULL);
    RegCloseKey(hKey);

    RegCloseKey(hKeyLM);
  }

  HKEY hKeyDefaultIcon;
  if (RegCreateKeyExW(
    HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW,
    0,
    NULL,
    REG_OPTION_NON_VOLATILE,
    KEY_WRITE,
    NULL,
    &hKeyDefaultIcon,
    NULL) == ERROR_SUCCESS
  ) {
    RegSetValueExW(
      hKeyDefaultIcon,
      L"DefaultIcon",
      0,
      REG_SZ,
      (BYTE*)L"%1",
      sizeof(L"%1"));
    RegCloseKey(hKeyDefaultIcon);
  }

  // Shell extension
  HKEY hKeyShellEx;
  if (RegCreateKeyExW(
    HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW L"\\ShellEx\\" CLSID_THUMBNAIL_HANDLER_GUID,
    0,
    NULL,
    REG_OPTION_NON_VOLATILE,
    KEY_WRITE,
    NULL,
    &hKeyShellEx,
    NULL) == ERROR_SUCCESS
  ) {
    RegSetValueExW(
      hKeyShellEx,
      NULL,
      0,
      REG_SZ,
      (BYTE*)THUMBNAIL_HANDLER_GUID,
      (wcslen(THUMBNAIL_HANDLER_GUID) + 1) * sizeof(WCHAR));
    RegCloseKey(hKeyShellEx);
  }

  // Add icon handler registration
  HKEY hKeyIconHandler;
  if (RegCreateKeyExW(
    HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW L"\\ShellEx\\IconHandler",
    0,
    NULL,
    REG_OPTION_NON_VOLATILE,
    KEY_WRITE,
    NULL,
    &hKeyIconHandler,
    NULL) == ERROR_SUCCESS
  ) {
    RegSetValueExW(
      hKeyIconHandler,
      NULL,
      0,
      REG_SZ,
      (BYTE*)THUMBNAIL_HANDLER_GUID,
      (wcslen(THUMBNAIL_HANDLER_GUID) + 1) * sizeof(WCHAR));
    RegCloseKey(hKeyIconHandler);
  }

  // Add thumbnail cutoff registration
  HKEY hKeyProgId;
  if (RegCreateKeyExW(
    HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW,
    0,
    NULL,
    REG_OPTION_NON_VOLATILE,
    KEY_WRITE,
    NULL,
    &hKeyProgId,
    NULL) == ERROR_SUCCESS
  ) {
    DWORD value = 0;
    RegSetValueExW(
      hKeyProgId,
      L"ThumbnailCutoff",
      0,
      REG_DWORD,
      reinterpret_cast<const BYTE*>(&value),
      sizeof(DWORD)
    );

    RegSetValueExW(
      hKeyProgId,
      L"Treatment",
      0,
      REG_DWORD,
      reinterpret_cast<const BYTE*>(&value),
      sizeof(DWORD)
    );
    RegCloseKey(hKeyProgId);
  }

  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

  return S_OK;
}

STDAPI DllUnregisterServer(void) {
  // Remove CLSID registration
  RegDeleteTreeW(HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\CLSID\\" THUMBNAIL_HANDLER_GUID);

  // Remove file associations
  RegDeleteTreeW(HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW L"\\ShellEx\\" CLSID_THUMBNAIL_HANDLER_GUID);

  // Remove icon handler registration
  RegDeleteTreeW(HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW L"\\ShellEx\\IconHandler");

  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

  return S_OK;
}
