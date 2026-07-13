// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#define _XLL_
#include <windows.h>
#include "xlcall.h"

typedef int (WINAPI *LPFNEXCEL12V)(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 opers[]);
static LPFNEXCEL12V pxExcel12v = NULL;

#define EXCEL12(fn, res, count, ...) \
    ((*pxExcel12v)((fn), (res), (count), __VA_ARGS__))

static DWORD   g_idCurrentSheet = 0;
static DWORD   g_idPrevSheet    = 0;
static BOOL    g_bInitialized   = FALSE;
static BOOL    g_bToggling      = FALSE;
static XCHAR   g_szCurrentBook[512] = {0};
static XCHAR   g_szPrevBook[512]    = {0};

static DWORD GetActiveSheetId(void)
{
    XLOPER12 xRes;
    EXCEL12(xlSheetId, &xRes, 1, TempActiveRef12());
    DWORD id = (xRes.xltype == xltypeNum) ? (DWORD)xRes.val.num : 0;
    EXCEL12(xlFree, 0, 1, &xRes);
    return id;
}

static void GetActiveWorkbookName(XCHAR *szBuffer, int bufSize)
{
    XLOPER12 xRes;
    szBuffer[0] = 0;
    EXCEL12(xlfGetDocument, &xRes, 1, TempInt12(88));
    if (xRes.xltype == xltypeStr && xRes.val.str != NULL) {
        int len = min(xRes.val.str[0], bufSize - 2);
        memcpy(szBuffer, xRes.val.str, (len + 1) * sizeof(XCHAR));
        szBuffer[len + 1] = 0;
    }
    EXCEL12(xlFree, 0, 1, &xRes);
}

static BOOL ActivateWorkbook(XCHAR *szBookName)
{
    if (szBookName[0] == 0) return FALSE;
    XLOPER12 xRes;
    EXCEL12(xlcWorkbookActivate, 0, 1, TempStr12(szBookName));
    return TRUE;
}

static void WINAPI ToggleToPreviousSheet(void)
{
    if (!g_bInitialized || g_szPrevBook[0] == 0) return;
    XCHAR szCurBook[512];
    DWORD curSheetId = GetActiveSheetId();
    GetActiveWorkbookName(szCurBook, sizeof(szCurBook));
    g_bToggling = TRUE;
    if (ActivateWorkbook(g_szPrevBook)) {
        XCHAR szTmpBook[512];
        memcpy(szTmpBook, g_szCurrentBook, sizeof(szTmpBook));
        memcpy(g_szCurrentBook, g_szPrevBook, sizeof(g_szCurrentBook));
        memcpy(g_szPrevBook, szCurBook, sizeof(g_szPrevBook));
        DWORD tmpSheet = g_idCurrentSheet;
        g_idCurrentSheet = g_idPrevSheet;
        g_idPrevSheet = curSheetId;
    }
    g_bToggling = FALSE;
}

static void WINAPI OnSheetActivate(void)
{
    if (g_bToggling) return;
    XCHAR szBook[512];
    DWORD sheetId = GetActiveSheetId();
    GetActiveWorkbookName(szBook, sizeof(szBook));
    if (!g_bInitialized) {
        memcpy(g_szCurrentBook, szBook, sizeof(g_szCurrentBook));
        g_idCurrentSheet = sheetId;
        g_bInitialized = TRUE;
        return;
    }
    if (memcmp(szBook, g_szCurrentBook, sizeof(szBook)) != 0 || sheetId != g_idCurrentSheet) {
        memcpy(g_szPrevBook, g_szCurrentBook, sizeof(g_szPrevBook));
        g_idPrevSheet = g_idCurrentSheet;
        memcpy(g_szCurrentBook, szBook, sizeof(g_szCurrentBook));
        g_idCurrentSheet = sheetId;
    }
}

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    XLOPER12 xDLL, xRegId;
    pxExcel12v = (LPFNEXCEL12V)GetProcAddress(GetModuleHandle(NULL), "Excel12v");
    if (!pxExcel12v) return 0;
    EXCEL12(xlGetName, &xDLL, 0);
    
    LPXLOPER12 args[7];
    XLOPER12 arg0, arg1, arg2, arg3, arg4, arg5, arg6;
    arg0.xltype = xltypeStr;  arg0.val.str = L"\x000B" L"ToggleSheet";
    arg1.xltype = xltypeStr;  arg1.val.str = L"\x0001" L"L";
    arg2.xltype = xltypeStr;  arg2.val.str = L"\x000B" L"ToggleSheet";
    arg3.xltype = xltypeStr;  arg3.val.str = L"\x0000";
    arg4.xltype = xltypeInt;  arg4.val.w = 1;
    arg5.xltype = xltypeStr;  arg5.val.str = L"\x000E" L"Previous Sheet";
    arg6.xltype = xltypeInt;  arg6.val.w = 0;
    args[0] = &arg0; args[1] = &arg1; args[2] = &arg2; args[3] = &arg3;
    args[4] = &arg4; args[5] = &arg5; args[6] = &arg6;
    
    EXCEL12(xlfRegister, &xRegId, 7, args[0], args[1], args[2], args[3], args[4], args[5], args[6]);
    EXCEL12(xlcOnKey, 0, 2, TempStr12(L"^%L"), TempStr12(L"ToggleSheet"));
    OnSheetActivate();
    return 1;
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    if (pxExcel12v) { EXCEL12(xlcOnKey, 0, 1, TempStr12(L"^%L")); }
    return 1;
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    if (fdwReason == DLL_PROCESS_ATTACH) { DisableThreadLibraryCalls(hinstDLL); }
    return TRUE;
}
