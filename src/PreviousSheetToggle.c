// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
// Uses low-level keyboard hook, no xlfRegister needed
#include <windows.h>

// Excel12v
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { const void *a[] = { __VA_ARGS__ }; pExcel12v((fn), (pxRes), (count), a); } while(0)

#define xltypeNum  0x0001
#define xltypeStr  0x0002
#define xltypeInt  0x0800
#define xltypeNil  0x0100

typedef struct {
    union { double num; wchar_t *str; unsigned short w; } val;
    unsigned short xltype;
} XLOPER12;

#define xlfGetDocument      20032
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

static BOOL    g_bHasPrev  = FALSE;
static BOOL    g_bToggling = FALSE;
static wchar_t g_prevBook[512];
static wchar_t g_prevSheet[512];

static wchar_t g_buf[4][256]; static int g_bi = 0;
static XLOPER12 Xls(const wchar_t *s) {
    wchar_t *b = g_buf[g_bi]; g_bi = (g_bi + 1) % 4;
    int len = (int)wcslen(s); if (len > 254) len = 254;
    b[0] = (wchar_t)len; memcpy(b + 1, s, len * 2);
    XLOPER12 x; x.xltype = xltypeStr; x.val.str = b; return x;
}
static XLOPER12 Xi(int i) {
    XLOPER12 x; x.xltype = xltypeInt; x.val.w = (unsigned short)i; return x;
}

static void Cur(wchar_t *bk, int bs, wchar_t *sh, int ss) {
    XLOPER12 r, a = Xi(88); EXCEL12(xlfGetDocument, &r, 1, &a);
    if (r.xltype == xltypeStr && r.val.str) {
        int len = r.val.str[0]; if (len > bs - 1) len = bs - 1;
        memcpy(bk, r.val.str + 1, len * 2); bk[len] = 0;
    } else bk[0] = 0;
    a = Xi(76); EXCEL12(xlfGetDocument, &r, 1, &a);
    if (r.xltype == xltypeStr && r.val.str) {
        int len = r.val.str[0]; if (len > ss - 1) len = ss - 1;
        memcpy(sh, r.val.str + 1, len * 2); sh[len] = 0;
    } else sh[0] = 0;
}

static void Go(const wchar_t *bk, const wchar_t *sh) {
    if (bk[0] == 0 || sh[0] == 0) return;
    XLOPER12 a = Xls(bk); EXCEL12(xlcWorkbookActivate, 0, 1, &a);
    a = Xls(sh); EXCEL12(xlcSelect, 0, 1, &a);
}

static void DoSwitch(void) {
    if (g_bToggling) return;
    wchar_t cb[512], cs[512]; Cur(cb, 512, cs, 512);
    if (!g_bHasPrev) {
        wcscpy(g_prevBook, cb); wcscpy(g_prevSheet, cs);
        g_bHasPrev = TRUE; return;
    }
    if (wcscmp(cb, g_prevBook) == 0 && wcscmp(cs, g_prevSheet) == 0) return;
    wchar_t ob[512], os[512];
    wcscpy(ob, g_prevBook); wcscpy(os, g_prevSheet);
    wcscpy(g_prevBook, cb); wcscpy(g_prevSheet, cs);
    g_bToggling = TRUE; Go(ob, os); g_bToggling = FALSE;
}

// ============================================================
// Keyboard hook
// ============================================================
static HHOOK g_hHook = NULL;
static HWND  g_hExcel = NULL;

static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)) {
        KBDLLHOOKSTRUCT *k = (KBDLLHOOKSTRUCT *)lParam;
        // Ctrl + Alt + L
        if (k->vkCode == 'L' &&
            (GetAsyncKeyState(VK_CONTROL) & 0x8000) &&
            (GetAsyncKeyState(VK_MENU) & 0x8000)) {
            DoSwitch();
            return 1; // Block further processing
        }
    }
    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

// ============================================================
// xlAutoOpen
// ============================================================
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    HMODULE h = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(h, "Excel12v");
    if (!pExcel12v) return 0;

    // Get Excel main window handle
    XLOPER12 r;
    EXCEL12(xlGetHwnd, &r, 0);
    g_hExcel = (r.xltype == xltypeNum) ? (HWND)(DWORD_PTR)r.val.num : GetActiveWindow();

    // Install keyboard hook on current thread
    g_hHook = SetWindowsHookEx(WH_KEYBOARD, KeyboardProc, NULL, GetCurrentThreadId());

    // Init state
    wchar_t b[512], s[512];
    Cur(b, 512, s, 512);
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
    if (g_hHook) {
        UnhookWindowsHookEx(g_hHook);
        g_hHook = NULL;
    }
    return 1;
}

// Need this
#define xlGetHwnd 20038

// ============================================================
// DllMain
// ============================================================
BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID p)
{
    if (r == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(h);
    return TRUE;
}
