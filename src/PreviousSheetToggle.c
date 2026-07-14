// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#include <windows.h>

// ============================================================
// Types and constants
// ============================================================
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

#define xltypeNum     0x0001
#define xltypeStr     0x0002
#define xltypeMissing 0x0080
#define xltypeNil     0x0100
#define xltypeInt     0x0800

typedef struct {
    union {
        double num;
        wchar_t *str;
        unsigned short w;
    } val;
    unsigned short xltype;
} XLOPER12;

#define xlfRegister         20001
#define xlfGetDocument      20032
#define xlcOnKey            30061
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

// ============================================================
// Global state
// ============================================================
static BOOL    g_bHasPrev  = FALSE;
static BOOL    g_bToggling = FALSE;
static wchar_t g_prevBook[512] = {0};
static wchar_t g_prevSheet[512] = {0};

// Static buffers for Excel strings
static wchar_t g_sb1[256], g_sb2[256], g_sb3[256], g_sb4[256];

// ============================================================
// Helper: build an Excel string (length-prefixed)
// ============================================================
static XLOPER12 Xls(const wchar_t *s, wchar_t *buf, int bufChars)
{
    int len = (int)wcslen(s);
    if (len > bufChars - 2) len = bufChars - 2;
    buf[0] = (wchar_t)len;
    if (len > 0) memcpy(buf + 1, s, len * sizeof(wchar_t));
    XLOPER12 x;
    x.xltype = xltypeStr;
    x.val.str = buf;
    return x;
}

static XLOPER12 Xi(int i)
{
    XLOPER12 x;
    x.xltype = xltypeInt;
    x.val.w = (unsigned short)i;
    return x;
}

// ============================================================
// Get current workbook and sheet names
// ============================================================
static void GetCurrent(wchar_t *book, int bsize, wchar_t *sheet, int ssize)
{
    XLOPER12 xArg, xRes;

    xArg = Xi(88);
    xRes.xltype = xltypeNil;
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > bsize - 1) len = bsize - 1;
        memcpy(book, xRes.val.str + 1, len * sizeof(wchar_t));
        book[len] = 0;
    } else {
        book[0] = 0;
    }

    xArg = Xi(76);
    xRes.xltype = xltypeNil;
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > ssize - 1) len = ssize - 1;
        memcpy(sheet, xRes.val.str + 1, len * sizeof(wchar_t));
        sheet[len] = 0;
    } else {
        sheet[0] = 0;
    }
}

// ============================================================
// Activate workbook and sheet
// ============================================================
static BOOL Activate(const wchar_t *book, const wchar_t *sheet)
{
    if (book[0] == 0 || sheet[0] == 0) return FALSE;

    XLOPER12 xb = Xls(book, g_sb3, 256);
    EXCEL12(xlcWorkbookActivate, 0, 1, &xb);

    XLOPER12 xs = Xls(sheet, g_sb4, 256);
    EXCEL12(xlcSelect, 0, 1, &xs);

    return TRUE;
}

// ============================================================
// Main command: record current and switch to previous
// ============================================================
__declspec(dllexport) void WINAPI SwitchCmd(void)
{
    if (g_bToggling) return;

    wchar_t curBook[512], curSheet[512];
    GetCurrent(curBook, 512, curSheet, 512);

    if (!g_bHasPrev) {
        // First call: just record current position
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
        g_bHasPrev = TRUE;
        return;
    }

    // Check if we're already on the "previous" sheet
    if (wcscmp(curBook, g_prevBook) == 0 && wcscmp(curSheet, g_prevSheet) == 0) {
        return; // Already there, nothing to do
    }

    // Save current as new previous
    wchar_t oldBook[512], oldSheet[512];
    wcscpy(oldBook, g_prevBook);
    wcscpy(oldSheet, g_prevSheet);
    wcscpy(g_prevBook, curBook);
    wcscpy(g_prevSheet, curSheet);

    // Jump to old position
    g_bToggling = TRUE;
    Activate(oldBook, oldSheet);
    g_bToggling = FALSE;
}

// ============================================================
// xlAutoOpen
// ============================================================
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    // Register command with shortcut "L"
    XLOPER12 xRes;
    XLOPER12 a0 = Xls(L"SwitchCmd", g_sb1, 256);
    XLOPER12 a1 = Xls(L"L", g_sb2, 256);

    static wchar_t nameBuf[32] = {9, L'S',L'w',L'i',L't',L'c',L'h',L'C',L'm',L'd'};
    XLOPER12 a2;
    a2.xltype = xltypeStr; a2.val.str = nameBuf;

    static wchar_t emptyBuf[2] = {0, 0};
    XLOPER12 a3;
    a3.xltype = xltypeStr; a3.val.str = emptyBuf;

    XLOPER12 a4 = Xi(1);

    static wchar_t catBuf[32] = {14, L'P',L'r',L'e',L'v',L'i',L'o',L'u',L's',L' ',L'S',L'h',L'e',L'e',L't'};
    XLOPER12 a5;
    a5.xltype = xltypeStr; a5.val.str = catBuf;

    XLOPER12 a6 = Xi(0);

    EXCEL12(xlfRegister, &xRes, 7, &a0, &a1, &a2, &a3, &a4, &a5, &a6);

    // Bind Ctrl+Alt+L
    static wchar_t keyBuf[8] = {4, L'^', L'%', L'L'};
    XLOPER12 xKey;
    xKey.xltype = xltypeStr; xKey.val.str = keyBuf;

    XLOPER12 xCmd;
    xCmd.xltype = xltypeStr; xCmd.val.str = nameBuf;

    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    // Init state
    wchar_t b[512], s[512];
    GetCurrent(b, 512, s, 512);
    wcscpy(g_prevBook, b);
    wcscpy(g_prevSheet, s);
    g_bHasPrev = TRUE;

    return 1;
}

// ============================================================
// xlAutoClose
// ============================================================
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    if (pExcel12v) {
        static wchar_t k[8] = {4, L'^', L'%', L'L'};
        XLOPER12 xKey;
        xKey.xltype = xltypeStr; xKey.val.str = k;
        EXCEL12(xlcOnKey, 0, 1, &xKey);
    }
    return 1;
}

// ============================================================
// DllMain
// ============================================================
BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID p)
{
    if (r == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(h);
    return TRUE;
}
