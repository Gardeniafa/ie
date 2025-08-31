// ie.cpp
// Build (no console window):
//   cl /EHsc /W4 /utf-8 ie.cpp ole32.lib oleaut32.lib /link /SUBSYSTEM:WINDOWS
//
// Example:
//   ie.exe
//   ie.exe https://a.com https://b.com
//   ie.exe --split https://a.com https://b.com

#define UNICODE
#define _UNICODE
#include <windows.h>
#include <shellapi.h>   // CommandLineToArgvW
#include <exdisp.h>     // IWebBrowser2
#include <exdispid.h>
#include <vector>
#include <string>
#include <cstdio>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

static bool g_console_attached = false;

static void attach_parent_console_if_any() {
    // If already has a console, do nothing.
    // Try to attach to parent console. If it fails, we stay GUI-only.
    if (AttachConsole(ATTACH_PARENT_PROCESS)) {
        g_console_attached = true;
        // Redirect C stdio to the attached console
        FILE* fp;
        (void)freopen_s(&fp, "CONOUT$", "w", stdout);
        (void)freopen_s(&fp, "CONOUT$", "w", stderr);
        (void)freopen_s(&fp, "CONIN$",  "r", stdin);
    }
}

static void show_error(HRESULT hr, const wchar_t* user_msg) {
    if (g_console_attached) {
        fwprintf(stderr, L"[ie] error: %s (hr=0x%08X)\n", user_msg, (unsigned)hr);
    } else {
        wchar_t buf[512];
        wsprintfW(buf, L"%s\nhr=0x%08X", user_msg, (unsigned)hr);
        MessageBoxW(NULL, buf, L"ie.exe", MB_OK | MB_ICONERROR | MB_TOPMOST);
    }
}

static void wait_until_idle(IWebBrowser2* pIE, DWORD max_ms) {
    if (!pIE) return;
    DWORD waited = 0;
    VARIANT_BOOL busy = VARIANT_FALSE;
    while (waited < max_ms && SUCCEEDED(pIE->get_Busy(&busy)) && busy == VARIANT_TRUE) {
        Sleep(50);
        waited += 50;
    }
}

static HRESULT open_urls_in_one_window(const std::vector<std::wstring>& urls) {
    if (urls.empty()) return E_INVALIDARG;

    IWebBrowser2* ie = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_InternetExplorer, nullptr, CLSCTX_LOCAL_SERVER,
                                  IID_IWebBrowser2, (void**)&ie);
    if (FAILED(hr)) {
        show_error(hr, L"failed to create InternetExplorer instance");
        return hr;
    }

    ie->put_Visible(VARIANT_TRUE);

    VARIANT vEmpty; VariantInit(&vEmpty);

    // First URL
    {
        VARIANT vURL; VariantInit(&vURL);
        vURL.vt = VT_BSTR;
        vURL.bstrVal = SysAllocString(urls[0].c_str());
        hr = ie->Navigate2(&vURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
        VariantClear(&vURL);
        if (FAILED(hr)) {
            show_error(hr, L"failed to navigate first URL");
            ie->Release();
            return hr;
        }
    }

    // Next URLs in new tabs
    for (size_t i = 1; i < urls.size(); ++i) {
        VARIANT vURL; VariantInit(&vURL);
        VARIANT vFlags; VariantInit(&vFlags);
        vURL.vt = VT_BSTR;
        vURL.bstrVal = SysAllocString(urls[i].c_str());
        vFlags.vt = VT_I4;
        vFlags.lVal = 0x800; // navOpenInNewTab

        hr = ie->Navigate2(&vURL, &vFlags, &vEmpty, &vEmpty, &vEmpty);
        VariantClear(&vURL);
        if (FAILED(hr)) {
            show_error(hr, L"failed to navigate next URL");
            ie->Release();
            return hr;
        }
    }

    wait_until_idle(ie, 2000);
    ie->Release();
    return S_OK;
}

static HRESULT open_urls_in_split_windows(const std::vector<std::wstring>& urls) {
    if (urls.empty()) return E_INVALIDARG;

    for (const auto& u : urls) {
        IWebBrowser2* ie = nullptr;
        HRESULT hr = CoCreateInstance(CLSID_InternetExplorer, nullptr, CLSCTX_LOCAL_SERVER,
                                      IID_IWebBrowser2, (void**)&ie);
        if (FAILED(hr)) {
            show_error(hr, L"failed to create InternetExplorer instance");
            return hr;
        }

        ie->put_Visible(VARIANT_TRUE);

        VARIANT vEmpty; VariantInit(&vEmpty);
        VARIANT vURL; VariantInit(&vURL);
        vURL.vt = VT_BSTR;
        vURL.bstrVal = SysAllocString(u.c_str());

        hr = ie->Navigate2(&vURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
        VariantClear(&vURL);
        if (FAILED(hr)) {
            show_error(hr, L"failed to navigate URL");
            ie->Release();
            return hr;
        }

        wait_until_idle(ie, 2000);
        ie->Release();
    }
    return S_OK;
}

// Entry point: Windows subsystem
int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int) {
    // Attach to parent console if launched from an existing shell.
    attach_parent_console_if_any();

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    std::vector<std::wstring> urls;
    bool split = false;

    if (argv && argc > 1) {
        for (int i = 1; i < argc; ++i) {
            std::wstring a = argv[i];
            if (a == L"--split") {
                split = true;
            } else {
                urls.push_back(a);
            }
        }
    }
    if (!argv) {
        show_error(E_FAIL, L"failed to parse command line");
    }

    if (urls.empty()) {
        urls.push_back(L"about:blank");
    }

    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        show_error(hr, L"CoInitialize failed");
        if (argv) LocalFree(argv);
        return 1;
    }

    if (split && urls.size() >= 2) {
        hr = open_urls_in_split_windows(urls);
    } else {
        hr = open_urls_in_one_window(urls);
    }

    if (FAILED(hr)) {
        // error already shown
        CoUninitialize();
        if (argv) LocalFree(argv);
        return 2;
    }

    CoUninitialize();
    if (argv) LocalFree(argv);
    return 0;
}
