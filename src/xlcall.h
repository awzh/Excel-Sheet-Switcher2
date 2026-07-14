// xlcall.h
#ifndef XLCALL_H
#define XLCALL_H

#ifdef __cplusplus
extern "C" {
#endif

typedef wchar_t XCHAR;

#define xltypeNum        0x0001
#define xltypeStr        0x0002
#define xltypeBool       0x0004
#define xltypeErr        0x0010
#define xltypeMissing    0x0080
#define xltypeNil        0x0100
#define xltypeInt        0x0800

typedef struct xloper12 {
    union {
        double num;
        wchar_t *str;
        WORD boolval;
        WORD err;
        short int w;
    } val;
    WORD xltype;
} XLOPER12;

extern int __stdcall Excel12v(int xlfn, void *pxResult, int count, const void *rgpx[]);

#define xlFree              10000
#define xlfRegister         20001
#define xlfGetWorkbook      20008
#define xlfGetDocument      20032
#define xlSheetId           20059
#define xlcOnKey            30061
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

#ifdef __cplusplus
}
#endif

#endif
