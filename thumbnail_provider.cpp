#include <atlbase.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <shlobj_core.h>
#include <shlwapi.h>
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
#define FILE_EXTENSION ".myextension"
#define FILE_EXTENSIONW L".myextension"

// Generate new UUID with `uuidgen -c` command in the terminal.
#define THUMBNAIL_HANDLER_GUID L"{972F8C22-1444-4075-A830-E7BA46C13FE9}"

#define CLSID_THUMBNAIL_HANDLER_GUID L"{E357FCCD-A995-4576-B01F-234630154E96}"

HMODULE g_hModule;

class __declspec(uuid("972F8C22-1444-4075-A830-E7BA46C13FE9")) ThumbnailProvider :
  public IInitializeWithFile,
  public IInitializeWithStream,
  public IThumbnailProvider {
public:
  IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
    static const QITAB qit[] = {
      QITABENT(ThumbnailProvider, IInitializeWithFile),
      QITABENT(ThumbnailProvider, IInitializeWithStream),
      QITABENT(ThumbnailProvider, IThumbnailProvider),
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

  // IInitializeWithStream
  IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode) {
    _pStream = pStream;

    return S_OK;
  }

  // IInitializeWithFile
  IFACEMETHODIMP Initialize(LPCWSTR pszFilePath, DWORD grfMode) {
    HRESULT hr = E_UNEXPECTED;
    if (_pszFilePath == NULL) {
      hr = StringCchCopyW(_szFilePath, ARRAYSIZE(_szFilePath), pszFilePath);

      if (SUCCEEDED(hr)) {
        _pszFilePath = _szFilePath;
      }
    }

    return hr;
  }

  std::vector<BYTE> ParseThumbNailFromFileData(const std::vector<BYTE>& fileData) {
    // Custom logic to extract thumbnail from file data as PNG.

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
        HANDLE hFile = CreateFileW(_pszFilePath, GENERIC_READ, FILE_SHARE_READ,
          NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
          DWORD fileSize = GetFileSize(hFile, NULL);
          fileData.resize(fileSize);
          DWORD bytesRead;
          ReadFile(hFile, fileData.data(), fileSize, &bytesRead, NULL);
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
  if (RegCreateKeyExW(HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\CLSID\\" THUMBNAIL_HANDLER_GUID,
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyLM, NULL) == ERROR_SUCCESS) {
    // InProcServer32
    HKEY hKeyServer;
    if (RegCreateKeyExW(hKeyLM, L"InProcServer32", 0, NULL,
      REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyServer, NULL) == ERROR_SUCCESS) {
      RegSetValueExW(hKeyServer, NULL, 0, REG_SZ,
        (BYTE*)szModule, (wcslen(szModule) + 1) * sizeof(WCHAR));
      RegSetValueExW(hKeyServer, L"ThreadingModel", 0, REG_SZ,
        (BYTE*)L"Apartment", sizeof(L"Apartment"));
      RegCloseKey(hKeyServer);
    }

    // Implemented Categories
    RegCreateKeyExW(hKeyLM, L"Implemented Categories\\" CLSID_THUMBNAIL_HANDLER_GUID,
      0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    RegCloseKey(hKey);

    RegCloseKey(hKeyLM);
  }

  // Shell extension
  HKEY hKeyShellEx;
  if (RegCreateKeyExW(HKEY_LOCAL_MACHINE,
    L"SOFTWARE\\Classes\\" FILE_EXTENSIONW L"\\ShellEx\\" CLSID_THUMBNAIL_HANDLER_GUID,
    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyShellEx, NULL) == ERROR_SUCCESS) {
    RegSetValueExW(hKeyShellEx, NULL, 0, REG_SZ,
      (BYTE*)THUMBNAIL_HANDLER_GUID,
      (wcslen(THUMBNAIL_HANDLER_GUID) + 1) * sizeof(WCHAR));
    RegCloseKey(hKeyShellEx);
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

  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

  return S_OK;
}
