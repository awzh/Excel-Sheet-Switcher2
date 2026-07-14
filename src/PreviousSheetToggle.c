// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#include <windows.h>
#include "xlcall.h"

// Excel12v function pointer
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

// Global state
static BOOL     g_bInitialized = FALSE;
static BOOL     g_bToggling    = FALSE;
static wchar_t  g_prevBook[512] = {0};
static wchar_t  g_prevSheet[512] = {0};

// ============================================================
// Helpers
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

// ============================================================
// Get current workbook name and sheet name
// ============================================================
static void GetCurrentInfo(wchar_t *bookName, int bookSize, wchar_t *sheetName, int sheetSize)
{
    XLOPER12 xRes;

    // Workbook name: xlfGetDocument(88)
    XLOPER12 xArg88 = MakeInt(88);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg88);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > bookSize - 1) len = bookSize - 1;
        memcpy(bookName, xRes.val.str + 1, len * sizeof(wchar_t));
        bookName[len] = 0;
    } else {
        bookName[0] = 0;
    }

    // Sheet name: xlfGetDocument(76)
    XLOPER12 xArg76 = MakeInt(76);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg76);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > sheetSize - 1) len = sheetSize - 1;
        memcpy(sheetName, xRes.val.str + 1, len * sizeof(wchar_t));
        sheetName[len] = 0;
    } else {
        sheetName[0] = 0;
    }
}

// ============================================================
// Activate a workbook and sheet
// ============================================================
static BOOL ActivateBookAndSheet(const wchar_t *bookName, const wchar_t *sheetName)
{
    if (bookName[0] == 0 || sheetName[0] == 0) return FALSE;

    // Build full reference string: "[BookName]SheetName"
    wchar_t fullRef[1024];
    wsprintf(fullRef, L"[%s]%s", bookName, sheetName);

    // Use xlcWorkbookActivate to activate workbook
    XLOPER12 xBook = MakeStr(bookName);
    EXCEL12(xlcWorkbookActivate, 0, 1, &xBook);

    // Use xlcSelect to jump to specific sheet
    XLOPER12 xRef = MakeStr(fullRef);
    EXCEL12(xlcSelect, 0, 1, &xRef);

    return TRUE;
}

// ============================================================
// Toggle command - this is what the shortcut calls
// ============================================================
static void WINAPI ToggleToPreviousSheet(void)
{
    if (!g_bInitialized) return;
    if (g_prevBook[0] == 0 || g_prevSheet[0] == 0) return;

    // Save current
    wchar_t curBook[512], curSheet[512];
    GetCurrentInfo(curBook, 512, curSheet, 512);

    g_bToggling = TRUE;

    if (ActivateBookAndSheet(g_prevBook, g_prevSheet)) {
        // Swap previous <-> current
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
    }

    g_bToggling = FALSE;
}

// ============================================================
// Called when user manually switches sheets
// ============================================================
static void UpdateHistory(void)
{
    if (g_bToggling) return;

    wchar_t curBook[512], curSheet[512];
    GetCurrentInfo(curBook, 512, curSheet, 512);

    if (!g_bInitialized) {
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
        g_bInitialized = TRUE;
        return;
    }

    // If changed, update previous
    if (wcscmp(curBook, g_prevBook) != 0 || wcscmp(curSheet, g_prevSheet) != 0) {
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
    }
}

// ============================================================
// Command wrapper that Excel calls (must be exported)
// ============================================================
__declspec(dllexport) void WINAPI ToggleCmd(void)
{
    // First update history (captures the sheet we're coming from)
    UpdateHistory();
    // Then toggle
    ToggleToPreviousSheet();
}

// ============================================================
// xlAutoOpen
// ============================================================
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    // Register command
    XLOPER12 xRes, xArgs[7];
    xArgs[0] = MakeStr(L"ToggleCmd");            // DLL export name
    xArgs[1] = MakeStr(L"L");                     // shortcut key
    xArgs[2] = MakeStr(L"ToggleCmd");            // Excel command name
    xArgs[3] = MakeStr(L"Switch to previous sheet"); // description
    xArgs[4] = MakeInt(1);                        // macro type: 1=command
    xArgs[5] = MakeStr(L"Previous Sheet");       // category
    xArgs[6] = MakeInt(0);                        // reserved

    EXCEL12(xlfRegister, &xRes, 7,
            &xArgs[0], &xArgs[1], &xArgs[2], &xArgs[3],
            &xArgs[4], &xArgs[5], &xArgs[6]);

    // Assign shortcut Ctrl+Alt+L
    XLOPER12 xKey = MakeStr(L"^%L");
    XLOPER12 xCmd = MakeStr(L"ToggleCmd");
    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    // Initialize history
    UpdateHistory();

    return 1;
}

// ============================================================
// xlAutoClose
// ============================================================
__declspec(dllexport) int WINAPI xlAutoClose(void)
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
