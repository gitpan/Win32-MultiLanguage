/* $Id$ */

/* max number of code pages */
#define MULTILANGUAGE_XS_NSCORES 32

/* :-( */
#define MyCoCreateMlang(p) \
    if (CoCreateInstance(&CLSID_CMultiLanguage, \
                         NULL, \
                         CLSCTX_ALL, \
                         &IID_IMultiLanguage2, \
                         (VOID**)&p) != S_OK) \
    { \
        warn("CoCreateInstance failed\n"); \
        XSRETURN_EMPTY; \
    }

#define MY_CLEANUP_AND_RETURN(p) \
        IMultiLanguage2_Release(p); \
        XSRETURN_EMPTY


#define VC_EXTRALEAN
#define CINTERFACE
#define COBJMACROS

#include <windows.h>
#include <mlang.h>

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"

SV* wchar2sv(LPCWSTR lpS, UINT nLen)
{
    LPSTR strBuf;
    int nRequired;
    SV* result;

    /* determine size for character buffer first */
    nRequired = WideCharToMultiByte(65001, 0, lpS, nLen, NULL, 0, NULL, NULL);
    
    if (!nRequired)
    {
        warn("Unexpected result from WideCharToMultiByte\n");
        return &PL_sv_undef;
    }

    New(42, strBuf, nRequired, char);
    
    if (!strBuf)
    {
        warn("Insufficient memory\n");
        return &PL_sv_undef;
    }
    
    if (!WideCharToMultiByte(65001, 0, lpS, nLen, strBuf, nRequired, NULL, NULL))
    {
        warn("WideCharToMultiByte failed\n");
        Safefree(strBuf);
        return &PL_sv_undef;
    }
    
    result = newSV(0);
    sv_usepvn(result, strBuf, nRequired);
    SvUTF8_on(result);

    return result;
}

MODULE = Win32::MultiLanguage       PACKAGE = Win32::MultiLanguage      

PROTOTYPES: DISABLE

BOOT:
    if (CoInitialize(NULL) != S_OK)
    {
        /* todo: check whether this is the best thing to do */
        croak("CoInitialize failed\n");
        XSRETURN_NO;
    }

void
DetectInputCodepage(octets, ...)
    SV* octets

  PREINIT:
    DWORD dwFlag = MLDETECTCP_NONE;
    DWORD dwPrefWinCodePage = 0;
    CHAR* pSrcStr;
    INT cSrcSize;
    DetectEncodingInfo lpEncoding[MULTILANGUAGE_XS_NSCORES];
    INT nScores = MULTILANGUAGE_XS_NSCORES;
    STRLEN nOctets;
    HRESULT hr;
    int i;
    AV* av;
    IMultiLanguage2* p;

  PPCODE:

    if (items > 1)
        dwFlag = (DWORD)SvIV(ST(1));

    if (items > 2)
        dwPrefWinCodePage = (DWORD)SvIV(ST(2));

    MyCoCreateMlang(p)

    pSrcStr = SvPV(octets, nOctets);
    cSrcSize = (INT)nOctets;

    /**/
    hr = IMultiLanguage2_DetectInputCodepage(p,
                                             dwFlag,
                                             dwPrefWinCodePage,
                                             pSrcStr,
                                             &cSrcSize,
                                             lpEncoding,
                                             &nScores);

    if (hr == S_FALSE)
    {
        MY_CLEANUP_AND_RETURN(p);
    }
    else if (hr == E_FAIL || hr != S_OK)
    {
        warn("An error occured while calling DetectInputCodepage\n");
        MY_CLEANUP_AND_RETURN(p);
    }
    
    av = newAV();

    for (i = 0; i < nScores; ++i)
    {
        HV* hv = newHV();

        hv_store(hv, "LangID", 6, newSViv(lpEncoding[i].nLangID), 0);
        hv_store(hv, "CodePage", 8, newSViv(lpEncoding[i].nCodePage), 0);
        hv_store(hv, "DocPercent", 10, newSViv(lpEncoding[i].nDocPercent), 0);
        hv_store(hv, "Confidence", 10, newSViv(lpEncoding[i].nConfidence), 0);

        av_push(av, newRV_noinc((SV*)hv));
    }

    /**/
    IMultiLanguage2_Release(p);
    XPUSHs(sv_2mortal(newRV_noinc((SV*)av)));

void
GetRfc1766FromLcid(svLocale)
    SV* svLocale

  PREINIT:
    BSTR bstrRfc1766;
    LCID lcidLocale;
    HRESULT hr;
    IMultiLanguage2* p;
    SV* result;

  PPCODE:
    lcidLocale = SvUV(svLocale);

    MyCoCreateMlang(p)
    
    hr = IMultiLanguage2_GetRfc1766FromLcid(p, lcidLocale, &bstrRfc1766);
    
    if (hr == E_INVALIDARG)
    {
        warn("One or more of the arguments are invalid.\n");
        MY_CLEANUP_AND_RETURN(p);
    }
    else if (hr == E_FAIL || hr != S_OK || !bstrRfc1766)
    {
        MY_CLEANUP_AND_RETURN(p);
    }
    
    result = wchar2sv(bstrRfc1766, wcslen(bstrRfc1766));
    
    /*
      it is not documented that the caller is responsible for freeing the bstr
      but if this is not done here the application leaks memory, so we free it
    */
    SysFreeString(bstrRfc1766);
    IMultiLanguage2_Release(p);
    XPUSHs(sv_2mortal(result));

void
GetCodePageInfo(svCodePage, svLangId)
    SV* svCodePage
    SV* svLangId

  PREINIT:
    UINT uiCodePage;
    LANGID LangId;
    MIMECPINFO cpi;
    IMultiLanguage2* p;
    HV* hv;
    HRESULT hr;

  PPCODE:
    uiCodePage = SvUV(svCodePage);
    LangId = SvUV(svLangId);

    MyCoCreateMlang(p)
    
    hr = IMultiLanguage2_GetCodePageInfo(p, uiCodePage, LangId, &cpi);

    if (hr != S_OK)
    {
        MY_CLEANUP_AND_RETURN(p);
    }
    
    hv = newHV();
    
    hv_store(hv, "Flags",             5, newSVuv(cpi.dwFlags),                                               0);
    hv_store(hv, "CodePage",          8, newSVuv(cpi.uiCodePage),                                            0);
    hv_store(hv, "FamilyCodePage",   14, newSVuv(cpi.uiFamilyCodePage),                                      0);
    hv_store(hv, "Description",      11, wchar2sv(cpi.wszDescription, wcslen(cpi.wszDescription)),           0);
    hv_store(hv, "WebCharset",       10, wchar2sv(cpi.wszWebCharset, wcslen(cpi.wszWebCharset)),             0);
    hv_store(hv, "HeaderCharset",    13, wchar2sv(cpi.wszHeaderCharset, wcslen(cpi.wszHeaderCharset)),       0);
    hv_store(hv, "BodyCharset",      11, wchar2sv(cpi.wszBodyCharset, wcslen(cpi.wszBodyCharset)),           0);
    hv_store(hv, "FixedWidthFont",   14, wchar2sv(cpi.wszFixedWidthFont, wcslen(cpi.wszFixedWidthFont)),     0);
    hv_store(hv, "ProportionalFont", 16, wchar2sv(cpi.wszProportionalFont, wcslen(cpi.wszProportionalFont)), 0);
    hv_store(hv, "GDICharset",       10, newSViv(cpi.bGDICharset),                                           0);

    IMultiLanguage2_Release(p);
    XPUSHs(sv_2mortal(newRV_noinc((SV*)hv)));

  /* todo: GetRfc1766Info */
  /* todo: GetFamilyCodePage */
  /* todo: GetCodePageDescription */
  /* todo: GetCharsetInfo */
  /* todo: DetectOutboundCodePage */

void
END()
  CODE:
    CoUninitialize();
    XSRETURN_YES; /* todo: ... */

