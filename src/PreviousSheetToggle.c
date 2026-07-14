#include <windows.h>

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
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
