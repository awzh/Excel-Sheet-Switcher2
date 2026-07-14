#include <windows.h>

__declspec(dllexport) void WINAPI ToggleCmd(void)
{
    MessageBoxW(NULL, L"ToggleCmd 执行成功！", L"Excel XLL 测试", MB_OK);
}

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    MessageBoxW(NULL, L"XLL 加载成功！", L"Excel XLL 测试", MB_OK);
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
