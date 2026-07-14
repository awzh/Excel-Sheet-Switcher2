// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#include <windows.h>
#include "xlcall.h"

// ============================================================
// Excel12v function pointer
// ============================================================
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

// Helper macro
#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

// ============================================================
// Global state
// ============================================================
static XLOPER12 g_xlCurrentSheet;
static XLOPER12 g_xlPrevSheet;
static BOOL     g_bInitialized = FALSE;
static BOOL     g_bToggling    = FALSE;

// ============================================================
// Helper: create string XLOPER12
// ============================================================
static XLOPER12 MakeStr(const wchar_t *s)
{
    XLOPER12 x;
    x.xltype = xltypeStr;
    x.val.str = (wchar_t*)s;
    return x;
}

static XLOPER12 MakeInt(int i)
{
    XLOPER12 x;
    x.xltype = xltypeInt;
    x.val.w = (WORD)i;
    return x;
}

static XLOPER12 MakeNum(double d)
{
    XLOPER12 x;
    x.xltype = xltypeNum;
    x.val.num = d;
    return x;
}

// ============================================================
// Core: Toggle to previous sheet
// ============================================================
static void WINAPI ToggleToPreviousSheet(void)
{
    if (!g_bInitialized) return;
    if (g_xlPrevSheet.xltype == xltypeNil) return;

    // Save current sheet info
    XLOPER12 xCur;
    xCur.xltype = xltypeNil;
    EXCEL12(xlSheetId, &xCur, 1, &g_xlMissing);

    // Activate previous sheet
    g_bToggling = TRUE;

    // Use xlcWorkbookActivate + xlcSheetActivate approach
    // Simplified: just use the previous sheet reference
    EXCEL12(xlSheetId, &g_xlMissing, 1, &g_xlPrevSheet);

    // Swap
    XLOPER12 xTmp = g_xlPrevSheet;
    g_xlPrevSheet = xCur;
    g_xlCurrentSheet = xTmp;

    g_bToggling = FALSE;
}

// ============================================================
// xlAutoOpen
// ============================================================
int WINAPI xlAutoOpen(void)
{
    // Get Excel12v
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    // Initialize globals
    g_xlMissing.xltype = xltypeMissing;
    g_xlCurrentSheet.xltype = xltypeNil;
    g_xlPrevSheet.xltype = xltypeNil;
    g_bInitialized = TRUE;

    // Register command
    XLOPER12 xResult, xArgs[7];

    xArgs[0] = MakeStr(L"ToggleSheet");              // procedure name
    xArgs[1] = MakeStr(L"L");                         // shortcut key
    xArgs[2] = MakeStr(L"ToggleSheet");              // command name
    xArgs[3] = MakeStr(L"");                          // description
    xArgs[4] = MakeInt(1);                            // macro type: command
    xArgs[5] = MakeStr(L"Previous Sheet");           // category
    xArgs[6] = MakeInt(0);                            // (reserved)

    EXCEL12(xlfRegister, &xResult, 7,
            &xArgs[0], &xArgs[1], &xArgs[2], &xArgs[3],
            &xArgs[4], &xArgs[5], &xArgs[6]);

    // Assign shortcut Ctrl+Alt+L
    XLOPER12 xKey = MakeStr(L"^%L");
    XLOPER12 xCmd = MakeStr(L"ToggleSheet");
    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    return 1;
}

// ============================================================
// xlAutoClose
// ============================================================
int WINAPI xlAutoClose(void)
{
    if (pExcel12v) {
        XLOPER12 xKey = MakeStr(L"^%L");
        EXCEL12(xlcOnKey, 0, 1, &xKey);
    }
    return 1;
}

// ============================================================
// DllMain
// ============================================================
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    if (fdwReason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(hinstDLL);
    }
    return TRUE;
}
