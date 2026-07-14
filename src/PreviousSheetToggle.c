#include <windows.h>

__declspec(dllexport) void WINAPI ToggleCmd(void)
{
    MessageBoxA(NULL, "ToggleCmd OK!", "XLL Test", MB_OK);
}

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    MessageBoxA(NULL, "XLL Loaded!", "XLL Test", MB_OK);
    return 1;
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    return 1;
}

BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID p)
{
    return TRUE;
}
