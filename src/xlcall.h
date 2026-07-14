// xlcall.h - Minimal Excel XLL SDK
#ifndef XLCALL_H
#define XLCALL_H

#ifdef __cplusplus
extern "C" {
#endif

// XLOPER12
#define xltypeNum  0x0001
#define xltypeStr  0x0002
#define xltypeInt  0x0800
#define xltypeNil  0x0100

typedef struct {
    union {
        double num;
        wchar_t *str;
        unsigned short w;
    } val;
    unsigned short xltype;
} XLOPER12;

typedef XLOPER12 *LPXLOPER12;

// Excel12v
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 opers[]);
extern LPFNEXCEL12V pxExcel12v;

#define Excel12v pxlcall_external_link  // This will be resolved at link time
// Actually we use the global pointer, so just use pxExcel12v directly

// Function indices
#define xlfRegister         20001
#define xlfGetDocument      20032
#define xlcOnKey            30061
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

#ifdef __cplusplus
}
#endif

#endif
