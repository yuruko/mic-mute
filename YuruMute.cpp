#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <CommCtrl.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <unordered_map>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Msimg32.lib")
#pragma comment(lib, "User32.lib")

// Control IDs for the Info Window
#define IDC_INFO_TEXT          1001
#define IDC_SELECT_DEVICE_BTN  1002
#define IDC_SET_HOTKEY_BTN     1003
#define IDC_INFO_LINK          1004

// Control IDs for the Hotkey Dialog
#define IDC_HOTKEY_LABEL       2001
#define IDC_HOTKEY_OK          2002
#define IDC_HOTKEY_CANCEL      2003

#define HOTKEY_ID 1
#define WM_TRAYICON (WM_USER + 1)

HINSTANCE hInst;
IAudioEndpointVolume* pEndpointVolume = NULL;
HWND hOverlayWnd = NULL;
HWND hInfoWnd = NULL; // Handle for the info window
BOOL bMuted = FALSE;
NOTIFYICONDATA nid = { 0 };
HMENU hTrayMenu = NULL;
std::wstring selectedDeviceName;
std::wstring configFilePath;

// Variables to store overlay position
int overlayPosX = 100; // Default X position
int overlayPosY = 100; // Default Y position

HICON hIconMicOn = NULL;  // Green icon
HICON hIconMicOff = NULL; // Red icon

int currentHotKey = VK_PAUSE;

bool isHotkeyDialogOpen = false;

// Mapping of special keys to readable names
std::unordered_map<UINT, std::wstring> specialKeyNames = {
    {VK_PAUSE, L"Pause"},
    {VK_F1, L"F1"},
    {VK_F2, L"F2"},
    {VK_F3, L"F3"},
    {VK_F4, L"F4"},
    {VK_F5, L"F5"},
    {VK_F6, L"F6"},
    {VK_F7, L"F7"},
    {VK_F8, L"F8"},
    {VK_F9, L"F9"},
    {VK_F10, L"F10"},
    {VK_F11, L"F11"},
    {VK_F12, L"F12"},
    {VK_F13, L"F13"},
    {VK_F14, L"F14"},
    {VK_F15, L"F15"},
    {VK_F16, L"F16"},
    {VK_F17, L"F17"},
    {VK_F18, L"F18"},
    {VK_F19, L"F19"},
    {VK_F20, L"F20"},
    {VK_F21, L"F21"},
    {VK_F22, L"F22"},
    {VK_F23, L"F23"},
    {VK_F24, L"F24"},
    {VK_HOME, L"Home"},
    {VK_END, L"End"},
    {VK_INSERT, L"Insert"},
    {VK_DELETE, L"Delete"},
    {VK_UP, L"Up Arrow"},
    {VK_DOWN, L"Down Arrow"},
    {VK_LEFT, L"Left Arrow"},
    {VK_RIGHT, L"Right Arrow"}
};

// Function Prototypes
void ToggleMute();
void UpdateOverlay();
void UpdateTrayIcon();
void LoadConfig();
void SaveConfig();
void SaveOverlayPosition();
void ShowDeviceSelectionDialog();
void InitializeAudioEndpoint();
void OpenInfoWindow();
void ShowHotkeyDialog(HWND hwnd);
HICON CreateColoredIcon(COLORREF color);
LRESULT CALLBACK OverlayWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK DeviceDialogProc(HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK InfoWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK HotkeyDialogProc(HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam);
std::wstring GetKeyName(UINT keyCode);
bool RegisterHotkeyDialogClass(HINSTANCE hInst);
void CreateTrayIcon(HWND hwnd);
void DestroyTrayIcon();
LRESULT CALLBACK LinkWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

// Main Function
int APIENTRY WinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPSTR lpCmdLine,
    _In_ int nCmdShow
)
{
    hInst = hInstance;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to initialize COM library.", L"Error", MB_OK | MB_ICONERROR);
        return 0;
    }

    // Load or initialize configuration
    PWSTR szPath = NULL;
    hr = SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, NULL, &szPath);
    if (SUCCEEDED(hr))
    {
        std::wstring configPath(szPath);
        CoTaskMemFree(szPath);

        configPath += L"\\YuruMute";
        CreateDirectory(configPath.c_str(), NULL);
        configPath += L"\\config.ini";
        configFilePath = configPath;
    }
    else
    {
        MessageBox(NULL, L"Failed to get AppData path.", L"Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return 0;
    }

    LoadConfig();
    InitializeAudioEndpoint();

    if (pEndpointVolume == NULL)
    {
        // Show the device selection dialog
        ShowDeviceSelectionDialog();

        // Try initializing again after device selection
        InitializeAudioEndpoint();

        if (pEndpointVolume == NULL)
        {
            MessageBox(NULL, L"No microphone device selected or device not found.", L"Error", MB_OK | MB_ICONERROR);
            CoUninitialize();
            return 0;
        }
    }

    pEndpointVolume->GetMute(&bMuted);

    // Create the red and green icons
    hIconMicOn = CreateColoredIcon(RGB(0, 255, 0));   // Green for Mic On
    hIconMicOff = CreateColoredIcon(RGB(255, 0, 0));  // Red for Mic Off

    if (!hIconMicOn || !hIconMicOff)
    {
        MessageBox(NULL, L"Failed to create icons.", L"Error", MB_OK | MB_ICONERROR);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    // Register the overlay window class
    WNDCLASS wcOverlay = { 0 };
    wcOverlay.lpfnWndProc = OverlayWndProc;
    wcOverlay.hInstance = hInstance;
    wcOverlay.lpszClassName = L"OverlayWindowClass";
    wcOverlay.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcOverlay.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    if (!RegisterClass(&wcOverlay)) {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to register OverlayWindowClass. Error code: " + std::to_wstring(error);
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        DestroyIcon(hIconMicOn);
        DestroyIcon(hIconMicOff);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    // Create the overlay window using saved position
    hOverlayWnd = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_NOACTIVATE,
        wcOverlay.lpszClassName, L"Overlay Window",
        WS_POPUP,
        overlayPosX, overlayPosY, 200, 50,
        NULL, NULL, hInstance, NULL);

    if (hOverlayWnd == NULL)
    {
        DWORD error = GetLastError();
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        std::wstring errorMessage = L"Failed to create overlay window. Error code: " + std::to_wstring(error);
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        DestroyIcon(hIconMicOn);
        DestroyIcon(hIconMicOff);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    SetLayeredWindowAttributes(hOverlayWnd, RGB(0, 0, 0), (BYTE)(255 * 0.7), LWA_ALPHA);

    ShowWindow(hOverlayWnd, SW_SHOWNOACTIVATE);
    UpdateOverlay();

    // Register the default hotkey
    if (!RegisterHotKey(NULL, HOTKEY_ID, 0, currentHotKey))
    {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to register hotkey. Error code: " + std::to_wstring(error);
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        DestroyWindow(hOverlayWnd);
        DestroyIcon(hIconMicOn);
        DestroyIcon(hIconMicOff);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    // Register the main application window class
    WNDCLASS wcApp = { 0 };
    wcApp.lpfnWndProc = WindowProc;
    wcApp.hInstance = hInstance;
    wcApp.lpszClassName = L"YuruMuteAppClass";

    if (!RegisterClass(&wcApp)) {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to register YuruMuteAppClass. Error code: " + std::to_wstring(error);
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        UnregisterClass(wcOverlay.lpszClassName, hInstance);
        DestroyWindow(hOverlayWnd);
        DestroyIcon(hIconMicOn);
        DestroyIcon(hIconMicOff);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    // Create a hidden window to handle messages
    HWND hWnd = CreateWindow(wcApp.lpszClassName, L"YuruMuteApp", 0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    if (!hWnd) {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to create main application window. Error code: " + std::to_wstring(error);
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        UnregisterClass(wcApp.lpszClassName, hInstance);
        UnregisterClass(wcOverlay.lpszClassName, hInstance);
        DestroyWindow(hOverlayWnd);
        DestroyIcon(hIconMicOn);
        DestroyIcon(hIconMicOff);
        pEndpointVolume->Release();
        CoUninitialize();
        return 0;
    }

    // Create the system tray icon
    CreateTrayIcon(hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        if (msg.message == WM_HOTKEY && msg.wParam == HOTKEY_ID)
        {
            ToggleMute();
            UpdateOverlay();
        }
        else if (msg.message == WM_APP) // Custom message to handle overlay update
        {
            UpdateOverlay();
        }

        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnregisterHotKey(NULL, HOTKEY_ID);
    DestroyTrayIcon();
    DestroyWindow(hOverlayWnd);
    UnregisterClass(wcApp.lpszClassName, hInstance);
    UnregisterClass(L"HotkeyDialogClass", hInstance);
    pEndpointVolume->Release();

    // Destroy icons
    if (hIconMicOn)
    {
        DestroyIcon(hIconMicOn);
        hIconMicOn = NULL;
    }
    if (hIconMicOff)
    {
        DestroyIcon(hIconMicOff);
        hIconMicOff = NULL;
    }

    CoUninitialize();

    return 0;
}

// Function to toggle mic mute
void ToggleMute()
{
    if (isHotkeyDialogOpen) {
        // Do not toggle mute while the hotkey dialog is open
        return;
    }

    pEndpointVolume->GetMute(&bMuted);
    pEndpointVolume->SetMute(!bMuted, NULL);
    bMuted = !bMuted;

    // Update tray icon
    UpdateTrayIcon();
}

// Function to update tray icon
void UpdateTrayIcon()
{
    nid.hIcon = bMuted ? hIconMicOff : hIconMicOn;
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

// Function to update the overlay
void UpdateOverlay()
{
    InvalidateRect(hOverlayWnd, NULL, TRUE);
    UpdateWindow(hOverlayWnd);
    UpdateTrayIcon(); // Update tray icon as well
}

// Function to load configuration
void LoadConfig()
{
    WCHAR buffer[256] = { 0 }; // Initialize buffer to zero
    GetPrivateProfileString(L"Settings", L"DeviceName", L"", buffer, 256, configFilePath.c_str());
    selectedDeviceName = buffer;

    // Load overlay position
    overlayPosX = GetPrivateProfileInt(L"Settings", L"OverlayPosX", 100, configFilePath.c_str());
    overlayPosY = GetPrivateProfileInt(L"Settings", L"OverlayPosY", 100, configFilePath.c_str());

    // Load hotkey
    WCHAR hotkeyBuffer[16] = { 0 };
    GetPrivateProfileString(L"Settings", L"HotKey", L"19", hotkeyBuffer, 16, configFilePath.c_str()); // Default to VK_PAUSE (0x13 == 19)
    currentHotKey = _wtoi(hotkeyBuffer);
}

// Function to save configuration
void SaveConfig()
{
    WritePrivateProfileString(L"Settings", L"DeviceName", selectedDeviceName.c_str(), configFilePath.c_str());
}

// Function to save overlay position
void SaveOverlayPosition()
{
    WCHAR buffer[16];
    wsprintf(buffer, L"%d", overlayPosX);
    WritePrivateProfileString(L"Settings", L"OverlayPosX", buffer, configFilePath.c_str());

    wsprintf(buffer, L"%d", overlayPosY);
    WritePrivateProfileString(L"Settings", L"OverlayPosY", buffer, configFilePath.c_str());
}

LRESULT CALLBACK DeviceDialogProc(HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static std::vector<std::wstring>* pDeviceNames;
    switch (msg)
    {
    case WM_CREATE:
        pDeviceNames = (std::vector<std::wstring>*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
        return 0;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_HOTKEY_OK)
        {
            HWND hList = GetDlgItem(hwndDlg, 1001);
            int sel = (int)SendMessageW(hList, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR)
            {
                WCHAR buffer[256] = { 0 };
                SendMessageW(hList, LB_GETTEXT, sel, (LPARAM)buffer);
                selectedDeviceName = buffer;
                SaveConfig();
                InitializeAudioEndpoint();
                UpdateOverlay();
            }
            DestroyWindow(hwndDlg);
            delete pDeviceNames;
            return 0;
        }
        else if (LOWORD(wParam) == IDC_HOTKEY_CANCEL)
        {
            DestroyWindow(hwndDlg);
            delete pDeviceNames;
            return 0;
        }
        break;
    case WM_DESTROY:
        return 0;
    }
    return DefWindowProc(hwndDlg, msg, wParam, lParam);
}

void ShowDeviceSelectionDialog()
{
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDeviceCollection* pCollection = NULL;

    HRESULT hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL,
        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        (void**)&pEnumerator);

    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to create MMDeviceEnumerator.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    hr = pEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pCollection);
    if (FAILED(hr))
    {
        MessageBox(NULL, L"Failed to enumerate audio endpoints.", L"Error", MB_OK | MB_ICONERROR);
        pEnumerator->Release();
        return;
    }

    UINT count;
    pCollection->GetCount(&count);

    std::vector<std::wstring>* pDeviceNames = new std::vector<std::wstring>();
    for (UINT i = 0; i < count; i++)
    {
        IMMDevice* pDevice = NULL;
        hr = pCollection->Item(i, &pDevice);
        if (SUCCEEDED(hr))
        {
            IPropertyStore* pProps = NULL;
            hr = pDevice->OpenPropertyStore(STGM_READ, &pProps);
            if (SUCCEEDED(hr))
            {
                PROPVARIANT varName;
                PropVariantInit(&varName);
                hr = pProps->GetValue(PKEY_Device_FriendlyName, &varName);
                if (SUCCEEDED(hr))
                {
                    pDeviceNames->push_back(varName.pwszVal);
                    PropVariantClear(&varName);
                }
                pProps->Release();
            }
            pDevice->Release();
        }
    }

    pCollection->Release();
    pEnumerator->Release();

    if (pDeviceNames->empty())
    {
        MessageBox(NULL, L"No audio capture devices found.", L"Error", MB_OK | MB_ICONERROR);
        delete pDeviceNames;
        return;
    }

    WNDCLASS wcDialog = { 0 };
    wcDialog.lpfnWndProc = DeviceDialogProc;
    wcDialog.hInstance = hInst;
    wcDialog.lpszClassName = L"SelectMicrophoneDialogClass";
    wcDialog.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcDialog.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    RegisterClass(&wcDialog);

    HWND hDialog = CreateWindowEx(
        WS_EX_DLGMODALFRAME,
        L"SelectMicrophoneDialogClass", L"select microphone",
        WS_VISIBLE | WS_SYSMENU | WS_CAPTION,
        CW_USEDEFAULT, CW_USEDEFAULT, 300, 146,
        NULL, NULL, hInst, NULL);

    if (!hDialog)
    {
        DWORD error = GetLastError();
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        std::wstring errorMessage = L"Failed to create device selection dialog. Error code: " + std::to_wstring(error);
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    SetWindowLongPtr(hDialog, GWLP_USERDATA, (LONG_PTR)pDeviceNames);

    RECT rcDlg, rcParent;
    GetWindowRect(hDialog, &rcDlg);
    GetWindowRect(GetDesktopWindow(), &rcParent);
    int x = (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
    int y = (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
    SetWindowPos(hDialog, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

    HWND hList = CreateWindowEx(WS_EX_CLIENTEDGE, L"LISTBOX", NULL,
        WS_CHILD | WS_VISIBLE | LBS_STANDARD,
        0, 0, 300, 100, hDialog, (HMENU)1001, hInst, NULL);

    for (size_t i = 0; i < pDeviceNames->size(); i++)
    {
        SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)pDeviceNames->at(i).c_str());
    }

    HWND hOkButton = CreateWindowW(L"BUTTON", L"ok",
        WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
        0, 83, 150, 25, hDialog, (HMENU)IDC_HOTKEY_OK, hInst, NULL);

    HWND hCancelButton = CreateWindowW(L"BUTTON", L"cancel",
        WS_CHILD | WS_VISIBLE,
        150, 83, 145, 25, hDialog, (HMENU)IDC_HOTKEY_CANCEL, hInst, NULL);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        if (!IsWindow(hDialog))
            break;
        if (msg.hwnd == hDialog || IsChild(hDialog, msg.hwnd))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        else
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
}

// Function to initialize the audio endpoint
void InitializeAudioEndpoint()
{
    if (pEndpointVolume)
    {
        pEndpointVolume->Release();
        pEndpointVolume = NULL;
    }

    if (selectedDeviceName.empty())
        return;

    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDeviceCollection* pCollection = NULL;

    HRESULT hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL,
        CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
        (void**)&pEnumerator);

    if (FAILED(hr))
    {
        return;
    }

    hr = pEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pCollection);
    if (FAILED(hr))
    {
        pEnumerator->Release();
        return;
    }

    UINT count;
    pCollection->GetCount(&count);

    for (UINT i = 0; i < count; i++)
    {
        IMMDevice* pDevice = NULL;
        hr = pCollection->Item(i, &pDevice);
        if (SUCCEEDED(hr))
        {
            IPropertyStore* pProps = NULL;
            hr = pDevice->OpenPropertyStore(STGM_READ, &pProps);
            if (SUCCEEDED(hr))
            {
                PROPVARIANT varName;
                PropVariantInit(&varName);
                hr = pProps->GetValue(PKEY_Device_FriendlyName, &varName);
                if (SUCCEEDED(hr))
                {
                    if (selectedDeviceName == varName.pwszVal)
                    {
                        hr = pDevice->Activate(__uuidof(IAudioEndpointVolume),
                            CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
                        PropVariantClear(&varName);
                        pProps->Release();
                        pDevice->Release();
                        break;
                    }
                    PropVariantClear(&varName);
                }
                pProps->Release();
            }
            pDevice->Release();
        }
    }

    pCollection->Release();
    pEnumerator->Release();
}

// Function to create the system tray icon
void CreateTrayIcon(HWND hwnd)
{
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hwnd;
    nid.uID = 1001;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;

    // Set the initial icon based on mute status
    nid.hIcon = bMuted ? hIconMicOff : hIconMicOn;
    wcscpy_s(nid.szTip, L"yuru mute mic");

    Shell_NotifyIcon(NIM_ADD, &nid);

    hTrayMenu = CreatePopupMenu();
    AppendMenuW(hTrayMenu, MF_STRING, 2000, L"yuru mute mic");
    AppendMenuW(hTrayMenu, MF_STRING, 2001, L"select device");
    AppendMenuW(hTrayMenu, MF_STRING, 2003, L"set hotkey"); // New Set Hotkey menu
    AppendMenuW(hTrayMenu, MF_STRING, 2002, L"exit");
}

// Function to destroy the system tray icon
void DestroyTrayIcon()
{
    Shell_NotifyIcon(NIM_DELETE, &nid);
    DestroyMenu(hTrayMenu);
}

// Function to open the Info window
void OpenInfoWindow()
{
    if (hInfoWnd != NULL && IsWindow(hInfoWnd))
    {
        // Bring the existing window to the foreground
        SetForegroundWindow(hInfoWnd);
        return;
    }

    // Define window class for the info window
    WNDCLASS wcInfo = { 0 };
    wcInfo.lpfnWndProc = InfoWndProc;
    wcInfo.hInstance = hInst;
    wcInfo.lpszClassName = L"YuruMuteInfoWindowClass";
    wcInfo.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcInfo.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1); // Use default window color

    // Register the class if not already registered
    WNDCLASS existingClass;
    if (!GetClassInfo(hInst, wcInfo.lpszClassName, &existingClass)) {
        if (!RegisterClass(&wcInfo)) {
            DWORD error = GetLastError();
            std::wstring errorMessage = L"Failed to register YuruMuteInfoWindowClass. Error code: " + std::to_wstring(error);
            MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
            return;
        }
    }

    // Create the info window
    hInfoWnd = CreateWindowEx(
        0,
        wcInfo.lpszClassName, L"yuru mute mic",
        WS_OVERLAPPED | WS_SYSMENU | WS_CAPTION,
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 98,
        NULL, NULL, hInst, NULL);

    if (!hInfoWnd)
    {
        DWORD error = GetLastError();
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        std::wstring errorMessage = L"Failed to create info window. Error code: " + std::to_wstring(error);
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Center the window
    RECT rcDlg, rcParent;
    GetWindowRect(hInfoWnd, &rcDlg);
    GetWindowRect(GetDesktopWindow(), &rcParent);
    int x = (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
    int y = (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
    SetWindowPos(hInfoWnd, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

    ShowWindow(hInfoWnd, SW_SHOW);
}

// Function to show the hotkey dialog
void ShowHotkeyDialog(HWND hwnd) {
    if (!RegisterHotkeyDialogClass(hInst)) {
        // Registration failed, do not proceed
        return;
    }

    // Set the flag to indicate the hotkey dialog is open
    isHotkeyDialogOpen = true;

    // Create a window for the hotkey dialog
    HWND hDlg = CreateWindowEx(
        0,                          // Optional window styles
        L"HotkeyDialogClass",       // Custom dialog class name
        L"hotkey",                  // Window title
        WS_OVERLAPPED | WS_VISIBLE, // Window style
        CW_USEDEFAULT, CW_USEDEFAULT,      // Position
        240, 103,                         // Width and height
        hwnd,                             // Parent window
        NULL,                             // Menu
        hInst,                            // Instance handle
        NULL                              // Additional application data
    );

    if (!hDlg) {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to create dialog. Error code: " + std::to_wstring(error);
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(hwnd, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        isHotkeyDialogOpen = false;  // Reset the flag if the dialog fails to open
        return;
    }

    // Show the dialog
    ShowWindow(hDlg, SW_SHOW);
}

// Function to register the custom dialog window class for Hotkey Dialog
bool RegisterHotkeyDialogClass(HINSTANCE hInst) {
    WNDCLASS wc = { 0 };

    wc.lpfnWndProc = HotkeyDialogProc;           // Window procedure for the dialog
    wc.hInstance = hInst;                      // Instance handle
    wc.lpszClassName = L"HotkeyDialogClass";       // Unique class name
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);  // Default window color

    // Check if the class is already registered
    WNDCLASS existingClass;
    if (GetClassInfo(hInst, wc.lpszClassName, &existingClass)) {
        // Class already exists, no need to register again
        return true;
    }

    // Register the class
    if (!RegisterClass(&wc)) {
        DWORD error = GetLastError();
        std::wstring errorMessage = L"Failed to register HotkeyDialogClass. Error code: " + std::to_wstring(error);
        LPWSTR errorMsg = NULL;
        FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPWSTR)&errorMsg,
            0,
            NULL
        );
        if (errorMsg) {
            errorMessage += L"\n" + std::wstring(errorMsg);
            LocalFree(errorMsg);
        }
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
        return false;
    }

    return true;
}

// Dialog procedure to handle key input and control actions
LRESULT CALLBACK HotkeyDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    static UINT newHotKey = VK_PAUSE;  // Default new hotkey

    switch (uMsg) {
    case WM_CREATE:
        // Create a label to display instructions
        CreateWindowW(L"STATIC", L"press any key to set as hotkey:", WS_VISIBLE | WS_CHILD,
            0, 0, 240, 20, hwndDlg, NULL, hInst, NULL);

        // Create a label to display the current hotkey name
        CreateWindowW(L"STATIC", GetKeyName(currentHotKey).c_str(), WS_VISIBLE | WS_CHILD,
            0, 20, 240, 20, hwndDlg, (HMENU)IDC_HOTKEY_LABEL, hInst, NULL);

        // Create OK button
        CreateWindowW(L"BUTTON", L"ok", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 40, 120, 25, hwndDlg, (HMENU)IDC_HOTKEY_OK, hInst, NULL);

        // Create Cancel button
        CreateWindowW(L"BUTTON", L"cancel", WS_VISIBLE | WS_CHILD,
            120, 40, 115, 25, hwndDlg, (HMENU)IDC_HOTKEY_CANCEL, hInst, NULL);
        return 0;

    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        // Capture the keypress and update the label
        newHotKey = (UINT)wParam;
        SetDlgItemTextW(hwndDlg, IDC_HOTKEY_LABEL, GetKeyName(newHotKey).c_str());
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_HOTKEY_OK:
            // Save the new hotkey and register it
            currentHotKey = newHotKey;
            UnregisterHotKey(NULL, HOTKEY_ID);
            if (!RegisterHotKey(NULL, HOTKEY_ID, 0, currentHotKey)) {
                MessageBox(hwndDlg, L"Failed to register new hotkey.", L"Error", MB_OK | MB_ICONERROR);
            }
            else {
                // Save the hotkey to the configuration file
                WCHAR buffer[16];
                wsprintf(buffer, L"%d", currentHotKey);
                WritePrivateProfileString(L"Settings", L"HotKey", buffer, configFilePath.c_str());
            }
            DestroyWindow(hwndDlg);
            return 0;

        case IDC_HOTKEY_CANCEL:
            DestroyWindow(hwndDlg);
            return 0;
        }
        break;

    case WM_CLOSE:
        DestroyWindow(hwndDlg);
        return 0;

    case WM_DESTROY:
        isHotkeyDialogOpen = false;  // Reset the flag when the dialog is closed
        return 0;
    }

    return DefWindowProc(hwndDlg, uMsg, wParam, lParam);
}

// Helper function to get the name of the key
std::wstring GetKeyName(UINT keyCode) {
    WCHAR name[128] = { 0 };

    // Check if it's a special key we need to handle manually
    if (specialKeyNames.find(keyCode) != specialKeyNames.end()) {
        return specialKeyNames[keyCode];
    }

    // Try using GetKeyNameText for non-special keys
    UINT scanCode = MapVirtualKey(keyCode, MAPVK_VK_TO_VSC);
    LONG lParamValue = (scanCode << 16);
    if (GetKeyNameTextW(lParamValue, name, 128) > 0) {
        return std::wstring(name);
    }

    // If no readable name was found, return a placeholder
    return L"Unknown Key";
}

// Overlay window procedure
LRESULT CALLBACK OverlayWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static POINT ptLast;
    switch (msg)
    {
    case WM_NCHITTEST:
    {
        // If CTRL is not pressed, make the window click-through
        if (!(GetKeyState(VK_CONTROL) & 0x8000))
        {
            return HTTRANSPARENT;
        }
        else
        {
            return HTCLIENT;
        }
    }

    case WM_LBUTTONDOWN:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            SetCapture(hwnd);
            GetCursorPos(&ptLast);
        }
        return 0;

    case WM_MOUSEMOVE:
        if ((wParam & MK_LBUTTON) && (GetKeyState(VK_CONTROL) & 0x8000))
        {
            POINT pt;
            GetCursorPos(&pt);
            int dx = pt.x - ptLast.x;
            int dy = pt.y - ptLast.y;
            RECT rc;
            GetWindowRect(hwnd, &rc);
            MoveWindow(hwnd, rc.left + dx, rc.top + dy, rc.right - rc.left, rc.bottom - rc.top, TRUE);
            ptLast = pt;
        }
        return 0;

    case WM_LBUTTONUP:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            ReleaseCapture();

            // Save the new position
            RECT rc;
            GetWindowRect(hwnd, &rc);
            overlayPosX = rc.left;
            overlayPosY = rc.top;
            SaveOverlayPosition();
        }
        return 0;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        RECT rect;
        GetClientRect(hwnd, &rect);

        HBRUSH hBrush = CreateSolidBrush(bMuted ? RGB(255, 0, 0) : RGB(0, 255, 0));
        FillRect(hdc, &rect, hBrush);
        DeleteObject(hBrush);

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, RGB(255, 255, 255));
        HFONT hFont = CreateFontW(24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS,
            ANTIALIASED_QUALITY, VARIABLE_PITCH, L"Arial");
        HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);

        LPCWSTR text = bMuted ? L"mic off" : L"mic on";

        DrawTextW(hdc, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        SelectObject(hdc, hOldFont);
        DeleteObject(hFont);

        EndPaint(hwnd, &ps);
    }
    return 0;

    case WM_DESTROY:
        // Do not call PostQuitMessage here
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// Info window procedure
LRESULT CALLBACK InfoWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        // Create static text with instructions
        CreateWindowW(L"STATIC", L"ctrl+drag to move overlay | yuru.be",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            0, 0, 400, 20, hwnd, NULL, hInst, NULL);

        // Create "Select Device" button
        HWND hSelectDeviceBtn = CreateWindowW(L"BUTTON", L"select device",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 20, 200, 40, hwnd, (HMENU)IDC_SELECT_DEVICE_BTN, hInst, NULL);

        // Create "Set Hotkey" button with unique ID
        HWND hSetHotkeyBtn = CreateWindowW(L"BUTTON", L"set hotkey",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            200, 20, 195, 40, hwnd, (HMENU)IDC_SET_HOTKEY_BTN, hInst, NULL);
        return 0;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_SELECT_DEVICE_BTN:
            ShowDeviceSelectionDialog();
            UpdateOverlay();
            return 0;

        case IDC_SET_HOTKEY_BTN:
            ShowHotkeyDialog(hwnd);
            return 0;
        }
        break;

    case WM_NOTIFY:
    {
        if (((LPNMHDR)lParam)->code == NM_CLICK || ((LPNMHDR)lParam)->code == NM_RETURN)
        {
            PNMLINK pNMLink = (PNMLINK)lParam;
            ShellExecute(NULL, L"open", pNMLink->item.szUrl, NULL, NULL, SW_SHOWNORMAL);
            return 0;
        }
        break;
    }

    case WM_DESTROY:
        hInfoWnd = NULL;
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// Function to create colored icons
HICON CreateColoredIcon(COLORREF color)
{
    int iconSize = GetSystemMetrics(SM_CXSMICON); // Get small icon size
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);

    HBITMAP hBitmapColor = CreateCompatibleBitmap(hdcScreen, iconSize, iconSize);
    HBITMAP hBitmapMask = CreateBitmap(iconSize, iconSize, 1, 1, NULL);

    HBITMAP hOldBitmap = (HBITMAP)SelectObject(hdcMem, hBitmapColor);

    // Fill the background with transparent color
    HBRUSH hBrushBackground = CreateSolidBrush(RGB(0, 0, 0));
    //FillRect(hdcMem, &(RECT){0, 0, iconSize, iconSize}, hBrushBackground);
    DeleteObject(hBrushBackground);

    // Draw a filled circle
    HBRUSH hBrush = CreateSolidBrush(color);
    SelectObject(hdcMem, hBrush);
    Ellipse(hdcMem, 0, 0, iconSize, iconSize);
    DeleteObject(hBrush);

    SelectObject(hdcMem, hOldBitmap);

    ICONINFO iconInfo = { 0 };
    iconInfo.fIcon = TRUE;
    iconInfo.hbmColor = hBitmapColor;
    iconInfo.hbmMask = hBitmapMask;

    HICON hIcon = CreateIconIndirect(&iconInfo);

    // Cleanup
    DeleteObject(hBitmapColor);
    DeleteObject(hBitmapMask);
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);

    return hIcon;
}

// Main application window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_TRAYICON)
    {
        if (lParam == WM_RBUTTONUP)
        {
            POINT pt;
            GetCursorPos(&pt);
            SetForegroundWindow(hwnd);
            TrackPopupMenu(hTrayMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
        }
        else if (lParam == WM_LBUTTONUP)
        {
            OpenInfoWindow();
        }
    }
    else if (msg == WM_COMMAND)
    {
        switch (LOWORD(wParam))
        {
        case 2000: // YuruMute
            OpenInfoWindow();
            return 0;

        case 2001: // Select Device
            ShowDeviceSelectionDialog();
            UpdateOverlay();
            return 0;

        case 2002: // Exit
            PostQuitMessage(0);
            return 0;

        case 2003: // Set Hotkey (from tray menu)
            ShowHotkeyDialog(hwnd);
            return 0;
        }
    }
    else if (msg == WM_DESTROY)
    {
        PostQuitMessage(0);
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// Link window procedure to handle hyperlink clicks
LRESULT CALLBACK LinkWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_MOUSEMOVE:
        // Let the default link behavior handle these messages
        break;
    case WM_NOTIFY:
    {
        LPNMHDR pnmh = (LPNMHDR)lParam;
        if (pnmh->code == NM_CLICK || pnmh->code == NM_RETURN)
        {
            PNMLINK pNMLink = (PNMLINK)lParam;
            ShellExecuteW(NULL, L"open", pNMLink->item.szUrl, NULL, NULL, SW_SHOWNORMAL);
            return 0;
        }
        break;
    }
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}
