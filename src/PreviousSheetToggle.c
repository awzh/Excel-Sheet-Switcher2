// PreviousSheetToggle.c - Excel XLL Sheet Switcher
#define _XLL_
#include <windows.h>
#include "xlcall.h"

// Global Excel12v function pointer
LPFNEXCEL12V pxExcel12v = NULL;

// Command ID returned by xlfRegister
static int g_cmdID = 0;

// State
static BOOL    g_bHasPrev  = FALSE;
static BOOL    g_bToggling = FALSE;
static wchar_t g_prevBook[512] = {0};
static wchar_t g_prevSheet[512] = {0};

// Get current workbook and sheet names
static void GetCurrent(wchar_t *book, int bsize, wchar_t *sheet, int ssize)
{
    XLOPER12 xRes, xArg;
    
    xArg.xltype = xltypeInt; xArg.val.w = 88;
    xRes.xltype = xltypeNil;
    Excel12v(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > bsize - 1) len = bsize - 1;
        memcpy(book, xRes.val.str + 1, len * sizeof(wchar_t));
        book[len] = 0;
    } else { book[0] = 0; }
    
    xArg.xltype = xltypeInt; xArg.val.w = 76;
    xRes.xltype = xltypeNil;
    Excel12v(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > ssize - 1) len = ssize - 1;
        memcpy(sheet, xRes.val.str + 1, len * sizeof(wchar_t));
        sheet[len] = 0;
    } else { sheet[0] = 0; }
}

// Activate workbook and sheet
static void Activate(const wchar_t *book, const wchar_t *sheet)
{
    if (book[0] == 0 || sheet[0] == 0) return;
    XLOPER12 x;
    x.xltype = xltypeStr;
    x.val.str = (wchar_t*)book;
    Excel12v(xlcWorkbookActivate, NULL, 1, &x);
    x.val.str = (wchar_t*)sheet;
    Excel12v(xlcSelect, NULL, 1, &x);
}

// The actual switch function - exported for OnKey
__declspec(dllexport) void WINAPI DoSwitch(void)
{
    if (g_bToggling) return;
    
    wchar_t cb[512], cs[512];
    GetCurrent(cb, 512, cs, 512);
    
    if (!g_bHasPrev) {
        wcscpy(g_prevBook, cb);
        wcscpy(g_prevSheet, cs);
        g_bHasPrev = TRUE;
        return;
    }
    
    if (wcscmp(cb, g_prevBook) == 0 && wcscmp(cs, g_prevSheet) == 0) return;
    
    wchar_t ob[512], os[512];
    wcscpy(ob, g_prevBook);
    wcscpy(os, g_prevSheet);
    wcscpy(g_prevBook, cb);
    wcscpy(g_prevSheet, cs);
    
    g_bToggling = TRUE;
    Activate(ob, os);
    g_bToggling = FALSE;
}

// xlAutoOpen
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xResult;
    
    // Get Excel12v
    pxExcel12v = (LPFNEXCEL12V)GetProcAddress(GetModuleHandle(NULL), "Excel12v");
    if (!pxExcel12v) return 0;
    
    // Register command using xlfRegister with proper string arguments
    // Arg0: Dll function name
    XLOPER12 arg0;
    arg0.xltype = xltypeStr;
    arg0.val.str = L"\x08DoSwitch";  // Length 8: DoSwitch
    
    // Arg1: type_text - " " means no built-in shortcut
    XLOPER12 arg1;
    arg1.xltype = xltypeStr;
    arg1.val.str = L"\x01 ";  // Length 1: space
    
    // Arg2: Function name in Excel
    XLOPER12 arg2;
    arg2.xltype = xltypeStr;
    arg2.val.str = L"\x08DoSwitch";
    
    // Arg3: Argument description - empty
    XLOPER12 arg3;
    arg3.xltype = xltypeStr;
    arg3.val.str = L"\x00";
    
    // Arg4: Macro type - 1 = command
    XLOPER12 arg4;
    arg4.xltype = xltypeInt;
    arg4.val.w = 1;
    
    // Arg5: Category
    XLOPER12 arg5;
    arg5.xltype = xltypeStr;
    arg5.val.str = L"\x0ESheet Switcher";  // Length 14
    
    // Arg6: Reserved
    XLOPER12 arg6;
    arg6.xltype = xltypeInt;
    arg6.val.w = 0;
    
    Excel12v(xlfRegister, &xResult, 7, &arg0, &arg1, &arg2, &arg3, &arg4, &arg5, &arg6);
    
    // Bind Ctrl+Alt+L
    XLOPER12 xKey, xCmd;
    xKey.xltype = xltypeStr;
    xKey.val.str = L"\x04^%L";  // Length 4: ^%L = Ctrl+Alt+L
    xCmd.xltype = xltypeStr;
    xCmd.val.str = L"\x08DoSwitch";
    Excel12v(xlcOnKey, NULL, 2, &xKey, &xCmd);
    
    // Init
    wchar_t b[512], s[512];
    GetCurrent(b, 512, s, 512);
    wcscpy(g_prevBook, b);
    wcscpy(g_prevSheet, s);
    g_bHasPrev = TRUE;
    
    return 1;
}

// xlAutoClose
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    if (pxExcel12v) {
        XLOPER12 xKey;
        xKey.xltype = xltypeStr;
        xKey.val.str = L"\x04^%L";
        Excel12v(xlcOnKey, NULL, 1, &xKey);
    }
    return 1;
}

// DllMain
BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID p)
{
    if (r == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(h);
    return TRUE;
}
