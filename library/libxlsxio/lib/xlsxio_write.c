#include "xlsxio_write.h"
#include "xlsxio_version.h"
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <time.h>
#include <inttypes.h>
#ifndef _WIN32
#include <unistd.h>
#endif
#include <fcntl.h>
#include <stdarg.h>

#ifdef USE_MINIZIP
#  include <minizip/zip.h>
#  if !defined(Z_DEFLATED) && defined(MZ_COMPRESS_METHOD_DEFLATE) /* support minizip2 which defines MZ_COMPRESS_METHOD_DEFLATE instead of Z_DEFLATED */
#    define Z_DEFLATED MZ_COMPRESS_METHOD_DEFLATE
#  endif
#  define ZIPFILETYPE zipFile
#else
#  if (defined(STATIC) || defined(BUILD_XLSXIO_STATIC) || defined(BUILD_XLSXIO_STATIC_DLL) || (defined(BUILD_XLSXIO) && !defined(BUILD_XLSXIO_DLL) && !defined(BUILD_XLSXIO_SHARED))) && !defined(ZIP_STATIC)
#    define ZIP_STATIC
#  endif
#  include <zip.h>
#  ifndef ZIP_RDONLY
typedef struct zip zip_t;
typedef struct zip_source zip_source_t;
#  endif
#  define ZIPFILETYPE zip_t
#  ifndef USE_LIBZIP
#    define USE_LIBZIP
#  endif
#endif

#if defined(_WIN32) && !defined(USE_PTHREADS)
#  define USE_WINTHREADS
#  include <windows.h>
#else
#  define USE_PTHREADS
#  include <pthread.h>
#endif

#if defined(_MSC_VER)
#  undef DLL_EXPORT_XLSXIO
#  define DLL_EXPORT_XLSXIO
#  define va_copy(dst,src) ((dst) = (src))
#endif

#ifdef _WIN32
#  define pipe(fds) _pipe(fds, 4096, _O_BINARY)
#  define read _read
#  define write _write
#  define write _write
#  define close _close
#  define fdopen _fdopen
#else
#  define _fdopen(f) f
#endif

//#undef WITHOUT_XLSX_STYLES
#define DEFAULT_BUFFERED_ROWS 5

#define FONT_CHAR_WIDTH 7
//#define CALCULATE_COLUMN_WIDTH(characters) ((double)characters + .75)
#define CALCULATE_COLUMN_WIDTH(characters) ((double)(long)(((long)characters * FONT_CHAR_WIDTH + 5) * 256 / FONT_CHAR_WIDTH) / 256.0)
#define CALCULATE_COLUMN_HEIGHT(characters) ((double)characters * 12.75)

DLL_EXPORT_XLSXIO void xlsxiowrite_get_version (int* pmajor, int* pminor, int* pmicro)
{
  if (pmajor)
    *pmajor = XLSXIO_VERSION_MAJOR;
  if (pminor)
    *pminor = XLSXIO_VERSION_MINOR;
  if (pmicro)
    *pmicro = XLSXIO_VERSION_MICRO;
}

DLL_EXPORT_XLSXIO const char* xlsxiowrite_get_version_string ()
{
  return XLSXIO_VERSION_STRING;
}

////////////////////////////////////////////////////////////////////////

//#define WITHOUT_XLSX_SHAREDSTRINGS
#define WITHOUT_XLSX_THEMES
//#define WITHOUT_XLSX_FOLDERS

//#define OPTIONAL_LINE_BREAK             "\r\n"
#define OPTIONAL_LINE_BREAK             ""
#define XML_HEADER                      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
#define XML_FOLDER_RELS                 "_rels/"
#ifndef WITHOUT_XLSX_FOLDERS
#define XML_FOLDER_DOCPROPS             "docProps/"
#define XML_FOLDER_XL                   "xl/"
#define XML_FOLDER_THEMES               "theme/"
#define XML_FOLDER_WORKSHEETS           "worksheets/"
#else
#define XML_FOLDER_DOCPROPS             ""
#define XML_FOLDER_XL                   ""
#define XML_FOLDER_WORKSHEETS           ""
#endif
#define XML_FILENAME_CONTENTTYPES       "[Content_Types].xml"
#define XML_FILENAME_RELS               ".rels"
#define XML_FILENAME_DOCPROPS_CORE      "core.xml"
#define XML_FILENAME_DOCPROPS_APP       "app.xml"
#define XML_FILENAME_XL_WORKBOOK_RELS   "workbook.xml.rels"
#define XML_FILENAME_XL_WORKBOOK        "workbook.xml"
#define XML_FILENAME_XL_STYLES          "styles.xml"
#define XML_FILENAME_XL_THEME1          "theme1.xml"
#define XML_FILENAME_XL_SHAREDSTRINGS   "sharedStrings.xml"
#define XML_FILENAME_XL_WORKSHEET1      "sheet1.xml"
#define XML_SHEETNAME_MAXLEN            31
#define XML_RELID_DOCPROPS_CORE             "rId2"
#define XML_RELID_DOCPROPS_APP              "rId3"
#define XML_RELID_XL_WORKBOOK               "rId1"
#define XML_WORKBOOK_RELID_XL_STYLES        "rId1"
#define XML_WORKBOOK_RELID_XL_WORKSHEET1    "rId2"
#define XML_WORKBOOK_RELID_XL_SHAREDSTRINGS "rId3"
#define XML_WORKBOOK_RELID_XL_THEME1        "rId4"

const char* content_types_xml =
  XML_HEADER
  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" OPTIONAL_LINE_BREAK
  "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" OPTIONAL_LINE_BREAK
  "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" OPTIONAL_LINE_BREAK
  "<Override PartName=\"/" XML_FOLDER_RELS XML_FILENAME_RELS "\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" OPTIONAL_LINE_BREAK
  "<Override PartName=\"/" XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" OPTIONAL_LINE_BREAK
  "<Override PartName=\"/" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE "\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" OPTIONAL_LINE_BREAK
  "<Override PartName=\"/" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP "\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" OPTIONAL_LINE_BREAK
#ifndef WITHOUT_XLSX_STYLES
  "<Override PartName=\"/" XML_FOLDER_XL XML_FILENAME_XL_STYLES "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" OPTIONAL_LINE_BREAK
#endif
#ifndef WITHOUT_XLSX_THEMES
  "<Override PartName=\"/" XML_FOLDER_XL XML_FOLDER_THEMES XML_FILENAME_XL_THEME1 "\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" OPTIONAL_LINE_BREAK
#endif
  "<Override PartName=\"/" XML_FOLDER_XL XML_FOLDER_RELS XML_FILENAME_XL_WORKBOOK_RELS "\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" OPTIONAL_LINE_BREAK
  "<Override PartName=\"/" XML_FOLDER_XL XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1 "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" OPTIONAL_LINE_BREAK
#ifndef WITHOUT_XLSX_SHAREDSTRINGS
  "<Override PartName=\"/" XML_FOLDER_XL XML_FILENAME_XL_SHAREDSTRINGS "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" OPTIONAL_LINE_BREAK
#endif
  "</Types>" OPTIONAL_LINE_BREAK;

const char* docprops_core_xml =
  XML_HEADER
  "<coreProperties xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">" OPTIONAL_LINE_BREAK
  //"<creator>" XLSXIOWRITE_FULLNAME "</creator>" OPTIONAL_LINE_BREAK
  "<lastModifiedBy>" XLSXIOWRITE_FULLNAME "</lastModifiedBy>" OPTIONAL_LINE_BREAK
  //"<modified>2016-04-24T17:50:35Z</modified>" OPTIONAL_LINE_BREAK
  "</coreProperties>" OPTIONAL_LINE_BREAK;

const char* docprops_app_xml =
  XML_HEADER
  "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" OPTIONAL_LINE_BREAK
  "<Application>" XLSXIOWRITE_NAME "</Application>" OPTIONAL_LINE_BREAK
  "<AppVersion>" XLSXIO_VERSION_STRING "</AppVersion>" OPTIONAL_LINE_BREAK
  "</Properties>" OPTIONAL_LINE_BREAK;

const char* rels_xml =
  XML_HEADER
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" OPTIONAL_LINE_BREAK
  "<Relationship Id=\"" XML_RELID_DOCPROPS_CORE "\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE "\"/>" OPTIONAL_LINE_BREAK
  "<Relationship Id=\"" XML_RELID_DOCPROPS_APP "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP "\"/>" OPTIONAL_LINE_BREAK
  "<Relationship Id=\"" XML_RELID_XL_WORKBOOK "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"" XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK "\"/>" OPTIONAL_LINE_BREAK
  "</Relationships>" OPTIONAL_LINE_BREAK;

#ifndef WITHOUT_XLSX_STYLES
const char* styles_xml =
  XML_HEADER
  "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" OPTIONAL_LINE_BREAK
  "<fonts count=\"2\">" OPTIONAL_LINE_BREAK
  "<font>" OPTIONAL_LINE_BREAK
  "<sz val=\"10\"/>" OPTIONAL_LINE_BREAK
  //"<color theme=\"1\"/>" OPTIONAL_LINE_BREAK
  "<name val=\"Consolas\"/>" OPTIONAL_LINE_BREAK
  "<family val=\"2\"/>" OPTIONAL_LINE_BREAK
  //"<scheme val=\"minor\"/>" OPTIONAL_LINE_BREAK
  "</font>" OPTIONAL_LINE_BREAK
  "<font>" OPTIONAL_LINE_BREAK
  "<b/><u/>"
  "<sz val=\"10\"/>" OPTIONAL_LINE_BREAK
  //"<color theme=\"1\"/>" OPTIONAL_LINE_BREAK
  "<name val=\"Consolas\"/>" OPTIONAL_LINE_BREAK
  "<family val=\"2\"/>" OPTIONAL_LINE_BREAK
  //"<scheme val=\"minor\"/>" OPTIONAL_LINE_BREAK
  "</font>" OPTIONAL_LINE_BREAK
  "</fonts>" OPTIONAL_LINE_BREAK
  "<fills count=\"1\">" OPTIONAL_LINE_BREAK
  "<fill/>" OPTIONAL_LINE_BREAK
  //"<fill><patternFill patternType=\"none\"/></fill>" OPTIONAL_LINE_BREAK
  "</fills>" OPTIONAL_LINE_BREAK
  "<borders count=\"2\">" OPTIONAL_LINE_BREAK
  "<border>" OPTIONAL_LINE_BREAK
  //"<left/>" OPTIONAL_LINE_BREAK
  //"<right/>" OPTIONAL_LINE_BREAK
  //"<top/>" OPTIONAL_LINE_BREAK
  //"<bottom/>" OPTIONAL_LINE_BREAK
  //"<diagonal/>" OPTIONAL_LINE_BREAK
  "</border>" OPTIONAL_LINE_BREAK
  "<border><bottom style=\"thin\"><color indexed=\"64\"/></bottom></border>" OPTIONAL_LINE_BREAK
  "</borders>" OPTIONAL_LINE_BREAK
  "<cellStyleXfs count=\"1\">" OPTIONAL_LINE_BREAK
  //"<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" OPTIONAL_LINE_BREAK
  "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">" OPTIONAL_LINE_BREAK
  "<alignment horizontal=\"general\" vertical=\"top\" wrapText=\"0\" shrinkToFit=\"0\" textRotation=\"0\" indent=\"0\"/>" OPTIONAL_LINE_BREAK
  "</xf>" OPTIONAL_LINE_BREAK
  "</cellStyleXfs>" OPTIONAL_LINE_BREAK
  "<cellXfs count=\"6\">" OPTIONAL_LINE_BREAK
  "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" OPTIONAL_LINE_BREAK
#define STYLE_HEADER 1
  "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>" OPTIONAL_LINE_BREAK
#define STYLE_GENERAL 2
  "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>" OPTIONAL_LINE_BREAK
#define STYLE_TEXT 3
  "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf>" OPTIONAL_LINE_BREAK
#define STYLE_INTEGER 4
  "<xf numFmtId=\"1\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>" OPTIONAL_LINE_BREAK
#define STYLE_DATETIME 5
  "<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"top\"/></xf>" OPTIONAL_LINE_BREAK
  "</cellXfs>" OPTIONAL_LINE_BREAK
  //"<cellStyles count=\"2\">" OPTIONAL_LINE_BREAK
  //"<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" OPTIONAL_LINE_BREAK
  //"</cellStyles>" OPTIONAL_LINE_BREAK
  "<dxfs count=\"0\"/>" OPTIONAL_LINE_BREAK
  //"<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>" OPTIONAL_LINE_BREAK
  "</styleSheet>" OPTIONAL_LINE_BREAK;
#endif

#ifndef WITHOUT_XLSX_THEMES
const char* theme_xml =
  XML_HEADER
  "<theme xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" OPTIONAL_LINE_BREAK
  "<themeElements><clrScheme name=\"Office\"><dk1><sysClr val=\"windowText\" lastClr=\"000000\"/></dk1><lt1><sysClr val=\"window\" lastClr=\"FFFFFF\"/></lt1><dk2><srgbClr val=\"1F497D\"/></dk2><lt2><srgbClr val=\"EEECE1\"/></lt2><accent1><srgbClr val=\"4F81BD\"/></accent1><accent2><srgbClr val=\"C0504D\"/></accent2><accent3><srgbClr val=\"9BBB59\"/></accent3><accent4><srgbClr val=\"8064A2\"/></accent4><accent5><srgbClr val=\"4BACC6\"/></accent5><accent6><srgbClr val=\"F79646\"/></accent6><hlink><srgbClr val=\"0000FF\"/></hlink><folHlink><srgbClr val=\"800080\"/></folHlink></clrScheme><fontScheme name=\"Office\"><majorFont><latin typeface=\"Cambria\"/><ea typeface=\"\"/><cs typeface=\"\"/><font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><font script=\"Hang\" typeface=\"맑은 고딕\"/><font script=\"Hans\" typeface=\"宋体\"/><font script=\"Hant\" typeface=\"新細明體\"/><font script=\"Arab\" typeface=\"Times New Roman\"/><font script=\"Hebr\" typeface=\"Times New Roman\"/><font script=\"Thai\" typeface=\"Tahoma\"/><font script=\"Ethi\" typeface=\"Nyala\"/><font script=\"Beng\" typeface=\"Vrinda\"/><font script=\"Gujr\" typeface=\"Shruti\"/><font script=\"Khmr\" typeface=\"MoolBoran\"/><font script=\"Knda\" typeface=\"Tunga\"/><font script=\"Guru\" typeface=\"Raavi\"/><font script=\"Cans\" typeface=\"Euphemia\"/><font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><font script=\"Thaa\" typeface=\"MV Boli\"/><font script=\"Deva\" typeface=\"Mangal\"/><font script=\"Telu\" typeface=\"Gautami\"/><font script=\"Taml\" typeface=\"Latha\"/><font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><font script=\"Orya\" typeface=\"Kalinga\"/><font script=\"Mlym\" typeface=\"Kartika\"/><font script=\"Laoo\" typeface=\"DokChampa\"/><font script=\"Sinh\" typeface=\"Iskoola Pota\"/><font script=\"Mong\" typeface=\"Mongolian Baiti\"/><font script=\"Viet\" typeface=\"Times New Roman\"/><font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></majorFont><minorFont><latin typeface=\"Calibri\"/><ea typeface=\"\"/><cs typeface=\"\"/><font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><font script=\"Hang\" typeface=\"맑은 고딕\"/><font script=\"Hans\" typeface=\"宋体\"/><font script=\"Hant\" typeface=\"新細明體\"/><font script=\"Arab\" typeface=\"Arial\"/><font script=\"Hebr\" typeface=\"Arial\"/><font script=\"Thai\" typeface=\"Tahoma\"/><font script=\"Ethi\" typeface=\"Nyala\"/><font script=\"Beng\" typeface=\"Vrinda\"/><font script=\"Gujr\" typeface=\"Shruti\"/><font script=\"Khmr\" typeface=\"DaunPenh\"/><font script=\"Knda\" typeface=\"Tunga\"/><font script=\"Guru\" typeface=\"Raavi\"/><font script=\"Cans\" typeface=\"Euphemia\"/><font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><font script=\"Thaa\" typeface=\"MV Boli\"/><font script=\"Deva\" typeface=\"Mangal\"/><font script=\"Telu\" typeface=\"Gautami\"/><font script=\"Taml\" typeface=\"Latha\"/><font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><font script=\"Orya\" typeface=\"Kalinga\"/><font script=\"Mlym\" typeface=\"Kartika\"/><font script=\"Laoo\" typeface=\"DokChampa\"/><font script=\"Sinh\" typeface=\"Iskoola Pota\"/><font script=\"Mong\" typeface=\"Mongolian Baiti\"/><font script=\"Viet\" typeface=\"Arial\"/><font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></minorFont></fontScheme><fmtScheme name=\"Office\"><fillStyleLst><solidFill><schemeClr val=\"phClr\"/></solidFill><gradFill rotWithShape=\"1\"><gsLst><gs pos=\"0\"><schemeClr val=\"phClr\"><tint val=\"50000\"/><satMod val=\"300000\"/></schemeClr></gs><gs pos=\"35000\"><schemeClr val=\"phClr\"><tint val=\"37000\"/><satMod val=\"300000\"/></schemeClr></gs><gs pos=\"100000\"><schemeClr val=\"phClr\"><tint val=\"15000\"/><satMod val=\"350000\"/></schemeClr></gs></gsLst><lin ang=\"16200000\" scaled=\"1\"/></gradFill><gradFill rotWithShape=\"1\"><gsLst><gs pos=\"0\"><schemeClr val=\"phClr\"><shade val=\"51000\"/><satMod val=\"130000\"/></schemeClr></gs><gs pos=\"80000\"><schemeClr val=\"phClr\"><shade val=\"93000\"/><satMod val=\"130000\"/></schemeClr></gs><gs pos=\"100000\"><schemeClr val=\"phClr\"><shade val=\"94000\"/><satMod val=\"135000\"/></schemeClr></gs></gsLst><lin ang=\"16200000\" scaled=\"0\"/></gradFill></fillStyleLst><lnStyleLst><ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><solidFill><schemeClr val=\"phClr\"><shade val=\"95000\"/><satMod val=\"105000\"/></schemeClr></solidFill><prstDash val=\"solid\"/></ln><ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><solidFill><schemeClr val=\"phClr\"/></solidFill><prstDash val=\"solid\"/></ln><ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><solidFill><schemeClr val=\"phClr\"/></solidFill><prstDash val=\"solid\"/></ln></lnStyleLst><effectStyleLst><effectStyle><effectLst><outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><srgbClr val=\"000000\"><alpha val=\"38000\"/></srgbClr></outerShdw></effectLst></effectStyle><effectStyle><effectLst><outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><srgbClr val=\"000000\"><alpha val=\"35000\"/></srgbClr></outerShdw></effectLst></effectStyle><effectStyle><effectLst><outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><srgbClr val=\"000000\"><alpha val=\"35000\"/></srgbClr></outerShdw></effectLst><scene3d><camera prst=\"orthographicFront\"><rot lat=\"0\" lon=\"0\" rev=\"0\"/></camera><lightRig rig=\"threePt\" dir=\"t\"><rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></lightRig></scene3d><sp3d><bevelT w=\"63500\" h=\"25400\"/></sp3d></effectStyle></effectStyleLst><bgFillStyleLst><solidFill><schemeClr val=\"phClr\"/></solidFill><gradFill rotWithShape=\"1\"><gsLst><gs pos=\"0\"><schemeClr val=\"phClr\"><tint val=\"40000\"/><satMod val=\"350000\"/></schemeClr></gs><gs pos=\"40000\"><schemeClr val=\"phClr\"><tint val=\"45000\"/><shade val=\"99000\"/><satMod val=\"350000\"/></schemeClr></gs><gs pos=\"100000\"><schemeClr val=\"phClr\"><shade val=\"20000\"/><satMod val=\"255000\"/></schemeClr></gs></gsLst><path path=\"circle\"><fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></path></gradFill><gradFill rotWithShape=\"1\"><gsLst><gs pos=\"0\"><schemeClr val=\"phClr\"><tint val=\"80000\"/><satMod val=\"300000\"/></schemeClr></gs><gs pos=\"100000\"><schemeClr val=\"phClr\"><shade val=\"30000\"/><satMod val=\"200000\"/></schemeClr></gs></gsLst><path path=\"circle\"><fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></path></gradFill></bgFillStyleLst></fmtScheme></themeElements><objectDefaults/><extraClrSchemeLst/>" OPTIONAL_LINE_BREAK
  "</theme>" OPTIONAL_LINE_BREAK;
#endif

#ifndef WITHOUT_XLSX_SHAREDSTRINGS
const char* sharedstrings_xml =
  XML_HEADER
  "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>" OPTIONAL_LINE_BREAK;
  //"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>" OPTIONAL_LINE_BREAK;
  //"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" OPTIONAL_LINE_BREAK
  //"<si><t></t></si>" OPTIONAL_LINE_BREAK
  //"</sst>" OPTIONAL_LINE_BREAK;
#endif

const char* workbook_rels_xml =
  XML_HEADER
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" OPTIONAL_LINE_BREAK
#ifndef WITHOUT_XLSX_STYLES
  "<Relationship Id=\"" XML_WORKBOOK_RELID_XL_STYLES "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"" XML_FILENAME_XL_STYLES "\"/>" OPTIONAL_LINE_BREAK
#endif
  "<Relationship Id=\"" XML_WORKBOOK_RELID_XL_WORKSHEET1 "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1 "\"/>" OPTIONAL_LINE_BREAK
#ifndef WITHOUT_XLSX_SHAREDSTRINGS
  "<Relationship Id=\"" XML_WORKBOOK_RELID_XL_SHAREDSTRINGS "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"" XML_FILENAME_XL_SHAREDSTRINGS "\"/>" OPTIONAL_LINE_BREAK
#endif
#ifndef WITHOUT_XLSX_THEMES
  "<Relationship Id=\"" XML_WORKBOOK_RELID_XL_THEME1 "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"" XML_FOLDER_THEMES XML_FILENAME_XL_THEME1 "\"/>" OPTIONAL_LINE_BREAK
#endif
  "</Relationships>" OPTIONAL_LINE_BREAK;

const char* workbook_xml =
  XML_HEADER
  "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" OPTIONAL_LINE_BREAK
  //"<workbookPr/>" OPTIONAL_LINE_BREAK
  "<bookViews>" OPTIONAL_LINE_BREAK
  //"<workbookView/>" OPTIONAL_LINE_BREAK
  "<workbookView activeTab=\"0\"/>" OPTIONAL_LINE_BREAK
  //"<workbookView activeTab=\"0\" xWindow=\"0\" yWindow=\"0\"/>" OPTIONAL_LINE_BREAK
  "</bookViews>" OPTIONAL_LINE_BREAK
  "<sheets>" OPTIONAL_LINE_BREAK
  "<sheet name=\"%s\" sheetId=\"1\" r:id=\"" XML_WORKBOOK_RELID_XL_WORKSHEET1 "\"/>" OPTIONAL_LINE_BREAK
  "</sheets>" OPTIONAL_LINE_BREAK
  "</workbook>" OPTIONAL_LINE_BREAK;

const char* worksheet_xml_begin =
  XML_HEADER
  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" OPTIONAL_LINE_BREAK;
  //"<dimension ref=\"A1\"/> OPTIONAL_LINE_BREAK"

const char* worksheet_xml_freeze_top_row =
  "<sheetViews>" OPTIONAL_LINE_BREAK
  "<sheetView tabSelected=\"1\" workbookViewId=\"0\">" OPTIONAL_LINE_BREAK
  "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>" OPTIONAL_LINE_BREAK
  "<selection pane=\"bottomLeft\"/>" OPTIONAL_LINE_BREAK
  "</sheetView>" OPTIONAL_LINE_BREAK
  "</sheetViews>" OPTIONAL_LINE_BREAK;

const char* worksheet_xml_start_data =
  "<sheetData>" OPTIONAL_LINE_BREAK;

const char* worksheet_xml_end =
  "</sheetData>" OPTIONAL_LINE_BREAK
  //"<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>" OPTIONAL_LINE_BREAK
  "</worksheet>" OPTIONAL_LINE_BREAK;

////////////////////////////////////////////////////////////////////////

#ifdef USE_LIBZIP
zip_int64_t zip_file_add_custom (ZIPFILETYPE* zip, const char* filename, zip_source_t* zipsrc)
{
  zip_int64_t index;
  if ((index = zip_file_add(zip, filename, zipsrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8)) >= 0) {
    zip_set_file_compression(zip, index, ZIP_CM_DEFLATE, 9);
    //zip_set_file_compression(zip, index, ZIP_CM_DEFLATE, 5);
    //zip_set_file_compression(zip, index, ZIP_CM_STORE, 0);
    //zip_file_set_external_attributes(zip, index, 0, ZIP_OPSYS_DOS, 0);
    zip_file_set_external_attributes(zip, index, 0, ZIP_OPSYS_VFAT, 0);
    //zip_file_set_comment(zip, index, "Test", 4, ZIP_FL_ENC_UTF_8);
    //zip_file_set_mtime(zip, index, time(NULL), 0);
    //zip_file_extra_field_delete(zip, index, ZIP_EXTRA_FIELD_ALL, ZIP_FL_CENTRAL | ZIP_FL_LOCAL);
  }
  return index;
}
#endif

int zip_add_content_buffer (ZIPFILETYPE* zip, const char* filename, const char* buf, size_t buflen, int mustfree)
{
#ifdef USE_MINIZIP
  zip_fileinfo zipinfo;
  time_t now = time(NULL);
  struct tm* newtm = localtime(&now);
  zipinfo.tmz_date.tm_sec = newtm->tm_sec;
  zipinfo.tmz_date.tm_min = newtm->tm_min;
  zipinfo.tmz_date.tm_hour = newtm->tm_hour;
  zipinfo.tmz_date.tm_mday = newtm->tm_mday;
  zipinfo.tmz_date.tm_mon = newtm->tm_mon;
  zipinfo.tmz_date.tm_year = newtm->tm_year;
  zipinfo.dosDate = 0;
  zipinfo.internal_fa = 0;
  zipinfo.external_fa = 0;
  if (zipOpenNewFileInZip(zip, filename, &zipinfo, NULL, 0, NULL, 0, NULL, Z_DEFLATED, 9) != ZIP_OK) {
    fprintf(stderr, "Error creating file \"%s\" inside zip file\n", filename);/////
    return 1;
  }
  if (zipWriteInFileInZip(zip, buf, buflen) != ZIP_OK) {
    fprintf(stderr, "Error writing to file \"%s\" inside zip file\n", filename);/////
    return 1;
  }
  zipCloseFileInZip(zip);
  if (mustfree)
    free((char*)buf);
#else
  zip_source_t* zipsrc;
  if ((zipsrc = zip_source_buffer(zip, buf, buflen, mustfree)) == NULL) {
    fprintf(stderr, "Error creating file \"%s\" inside zip file\n", filename);/////
    return 1;
  }
  if (zip_file_add_custom(zip, filename, zipsrc) < 0) {
    fprintf(stderr, "Error in zip_file_add for file %s\n", filename);/////
    zip_source_free(zipsrc);
    return 2;
  }
#endif
  return 0;
}

int zip_add_static_content_string (ZIPFILETYPE* zip, const char* filename, const char* data)
{
  return zip_add_content_buffer(zip, filename, data, strlen(data), 0);
}

int zip_add_dynamic_content_string (ZIPFILETYPE* zip, const char* filename, const char* data, ...)
{
  int result;
  char* buf;
  int buflen;
  va_list args;
  va_start(args, data);
  buflen = vsnprintf(NULL, 0, data, args);
  if (buflen < 0 || (buf = (char*)malloc(buflen + 1)) == NULL) {
    result = -1;
  } else {
    va_end(args);
    va_start(args, data);
    vsnprintf(buf, buflen + 1, data, args);
    result = zip_add_content_buffer(zip, filename, buf, buflen, 1);
  }
  va_end(args);
  return result;
}

////////////////////////////////////////////////////////////////////////

//replace part of a string
char* str_replace (char** s, size_t pos, size_t len, char* replacement)
{
  if (!s || !*s)
    return NULL;
  size_t totallen = strlen(*s);
  size_t replacementlen = strlen(replacement);
  if (pos > totallen)
    pos = totallen;
  if (pos + len > totallen)
    len = totallen - pos;
  if (replacementlen > len)
    if ((*s = (char*)realloc(*s, totallen - len + replacementlen + 1)) == NULL)
      return NULL;
  memmove(*s + pos + replacementlen, *s + pos + len, totallen - pos - len + 1);
  memcpy(*s + pos, replacement, replacementlen);
  return *s;
}

//fix string for use as XML data
char* fix_xml_special_chars (char** s)
{
	int pos = 0;
	while (*s && (*s)[pos]) {
		switch ((*s)[pos]) {
			case '&' :
				str_replace(s, pos, 1, "&amp;");
				pos += 5;
				break;
			case '\"' :
				str_replace(s, pos, 1, "&quot;");
				pos += 6;
				break;
			case '\'' :
				str_replace(s, pos, 1, "&apos;");
				pos += 6;
				break;
			case '<' :
				str_replace(s, pos, 1, "&lt;");
				pos += 4;
				break;
			case '>' :
				str_replace(s, pos, 1, "&gt;");
				pos += 4;
				break;
			case '\r' :
				str_replace(s, pos, 1, "");
				break;
			default:
				pos++;
				break;
		}
	}
	return *s;
}

/*
//add data to a null-terminated buffer and update the length counter
int vappend_data (char** pdata, size_t* pdatalen, const char* format, va_list args)
{
  int len;
  va_list args2;
  va_copy(args2, args);
  //va_start(args, format);
  if ((len = vsnprintf(NULL, 0, format, args)) < 0)
    return -1;
  va_end(args);
  if ((*pdata = (char*)realloc(*pdata, *pdatalen + len + 1)) == NULL)
    return -1;
  vsnprintf(*pdata + *pdatalen, len + 1, format, args2);
  va_end(args2);
  *pdatalen += len;
  return len;
}

//add formatted data to a null-terminated buffer and update the length counter
int append_data (char** pdata, size_t* pdatalen, const char* format, ...)
{
  int result;
  va_list args;
  va_start(args, format);
  result = vappend_data(pdata, pdatalen, format, args);
  va_end(args);
  return result;
}
*/

//add formatted data to a null-terminated buffer and update the length counter
int append_data (char** pdata, size_t* pdatalen, const char* format, ...)
{
  int len;
  va_list args;
  va_start(args, format);
  len = vsnprintf(NULL, 0, format, args);
  va_end(args);
  if (len < 0)
    return -1;
  if ((*pdata = (char*)realloc(*pdata, *pdatalen + len + 1)) == NULL)
    return -1;
  va_start(args, format);
  vsnprintf(*pdata + *pdatalen, len + 1, format, args);
  va_end(args);
  *pdatalen += len;
  return len;
}

#ifndef NO_COLUMN_NUMBERS
/*
//insert formatted data into a null-terminated buffer at the specified position and update the length counter
int insert_data (char** pdata, size_t* pdatalen, size_t pos, const char* format, ...)
{
  int len;
  va_list args;
  va_start(args, format);
  len = vsnprintf(NULL, 0, format, args);
  va_end(args);
  if (len < 0)
    return -1;
  if ((*pdata = (char*)realloc(*pdata, *pdatalen + len + 1)) == NULL)
    return -1;
  if (pos > *pdatalen)
    pos = *pdatalen;
  if (pos < *pdatalen)
    memmove(*pdata + pos + len, *pdata + pos, *pdatalen - pos + 1);
  else
    (*pdata)[pos + len] = 0;
  va_start(args, format);
  len = vsnprintf(*pdata + pos, len, format, args);
  va_end(args);
  *pdatalen += len;
  return len;
}

char* get_A1col (uint64_t col)
{
  char* result = NULL;
  size_t resultlen = 0;
  if (col > 0) {
    do {
      col--;
      insert_data(&result, &resultlen, 0, "%c", 'A' + (col % 26));
      col = col / 26;
    } while (col > 0);
  }
  return result;
}
*/

char* get_A1col (uint64_t col)
{
  char* result = NULL;
  size_t resultlen = 0;
  //allocate 19 bytes as the maximum value for 64-bit devided by 26 has 18 digits
  if (col > 0 && (result = (char*)malloc(19)) != NULL) {
    result[0] = 0;
    do {
      col--;
      memmove(result + 1, result, ++resultlen);
      result[0] = 'A' + (col % 26);
      col = col / 26;
    } while (col > 0);
  }
  return result;
}
#endif

#define need_space_preserve_attr(value) 1
/*
int need_space_preserve_attr (const char* value)
{
  /////TO DO: return non-zero only if space at beginning or end, or contains multiple consecutive spaces
  return 1;
}
*/

////////////////////////////////////////////////////////////////////////

struct column_info_struct {
  int width;
  int maxwidth;
  struct column_info_struct* next;
};

struct xlsxio_write_struct {
  char* filename;
  char* sheetname;
  ZIPFILETYPE* zip;
#ifdef USE_WINTHREADS
  HANDLE thread;
#else
  pthread_t thread;
#endif
  FILE* pipe_read;
  FILE* pipe_write;
  struct column_info_struct* columninfo;
  struct column_info_struct** pcurrentcolumn;
  char* buf;
  size_t buflen;
  size_t rowstobuffer;
  size_t rowheight;
  int freezetop;
  int sheetopen;
  int rowopen;
#ifndef NO_ROW_NUMBERS
  uint64_t rownr;
#ifndef NO_COLUMN_NUMBERS
  uint64_t colnr;
#endif
#endif
};

#ifndef NO_ROW_NUMBERS
#define ROWNRTAG " r=\"%" PRIu64 "\""
#define ROWNRPARAM(handle) , handle->rownr
#ifndef NO_COLUMN_NUMBERS
#define COLNRTAG " r=\"%s%" PRIu64 "\""
#endif
#endif

//thread function used for creating .xlsx file from pipe
#ifdef USE_WINTHREADS
DWORD WINAPI thread_proc (LPVOID arg)
#else
void* thread_proc (void* arg)
#endif
{
  xlsxiowriter handle = (xlsxiowriter)arg;
  //generate required files
  zip_add_static_content_string(handle->zip, XML_FILENAME_CONTENTTYPES, content_types_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE, docprops_core_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP, docprops_app_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_RELS XML_FILENAME_RELS, rels_xml);
#ifndef WITHOUT_XLSX_STYLES
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_STYLES, styles_xml);
#endif
#ifndef WITHOUT_XLSX_THEMES
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FOLDER_THEMES XML_FILENAME_XL_THEME1, theme_xml);
#endif
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FOLDER_RELS XML_FILENAME_XL_WORKBOOK_RELS, workbook_rels_xml);
  { //TO DO: this crashes on Linux
    char* sheetname = NULL;
    if (handle->sheetname) {
      sheetname = strdup(handle->sheetname);
      if (sheetname) {
        if (strlen(sheetname) > XML_SHEETNAME_MAXLEN)
          sheetname[XML_SHEETNAME_MAXLEN] = 0;
        fix_xml_special_chars(&sheetname);
      }
    }
    zip_add_dynamic_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK, workbook_xml, (sheetname ? sheetname : "Sheet1"));
    free(sheetname);
  }
#ifndef WITHOUT_XLSX_SHAREDSTRINGS
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_SHAREDSTRINGS, sharedstrings_xml);
#endif

  //add sheet content file with pipe data
#ifdef USE_MINIZIP
#define MINIZIP_PIPE_BUFFER_SIZE 1024
//#error TO DO:
  if (zipOpenNewFileInZip(handle->zip, XML_FOLDER_XL XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1, NULL, NULL, 0, NULL, 0, NULL, Z_DEFLATED, 9) != ZIP_OK) {
    fprintf(stderr, "Error adding file");
  } else {
    char* buf;
    size_t buflen;
    if ((buf = (char*)malloc(MINIZIP_PIPE_BUFFER_SIZE)) == NULL) {
      fprintf(stderr, "Memory allocation error");/////
    } else {
      while ((buflen = fread(buf, 1, MINIZIP_PIPE_BUFFER_SIZE, handle->pipe_read)) > 0) {
        if (zipWriteInFileInZip(handle->zip, buf, buflen) != ZIP_OK) {
          fprintf(stderr, "Error writing file inside archive");/////
          break;
        }
      }
      free(buf);
    }
    fclose(handle->pipe_read);
    zipCloseFileInZip(handle->zip);
    zipClose(handle->zip, NULL);
  }
#else
  zip_source_t* zipsrc = zip_source_filep(handle->zip, handle->pipe_read, 0, -1);
  if (zip_file_add_custom(handle->zip, XML_FOLDER_XL XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1, zipsrc) < 0) {
    zip_source_free(zipsrc);
    fprintf(stderr, "Error adding file");
  }
#ifdef ZIP_RDONLY
  zip_file_set_mtime(handle->zip, zip_get_num_entries(handle->zip, 0) - 1, time(NULL), 0);
#endif
  //close zip file (processes all data, will block until pipe is closed)
  if (zip_close(handle->zip) != 0) {
    int ze, se;
#ifdef ZIP_RDONLY
    zip_error_t* error = zip_get_error(handle->zip);
    ze = zip_error_code_zip(error);
    se = zip_error_code_system(error);
#else
    zip_error_get(handle->zip, &ze, &se);
#endif
    fprintf(stderr, "zip_close failed (%i,%i)\n", ze, se);/////
    fprintf(stderr, "can't close zip archive : %s\n", zip_strerror(handle->zip));
  }
#endif
  handle->zip = NULL;
  handle->pipe_read = NULL;
#ifdef USE_WINTHREADS
  return 0;
#else
  return NULL;
#endif
}

////////////////////////////////////////////////////////////////////////

DLL_EXPORT_XLSXIO xlsxiowriter xlsxiowrite_open (const char* filename, const char* sheetname)
{
  xlsxiowriter handle;
  if (!filename)
    return NULL;
  if ((handle = (xlsxiowriter)malloc(sizeof(struct xlsxio_write_struct))) != NULL) {
    int pipefd[2];
    //initialize
    handle->filename = strdup(filename);
    handle->sheetname = (sheetname ? strdup(sheetname) : NULL);
    handle->zip = NULL;
    //handle->pipe_read = NULL;
    //handle->pipe_write = NULL;
    handle->columninfo = NULL;
    handle->pcurrentcolumn = &handle->columninfo;
    handle->buf = NULL;
    handle->buflen = 0;
    handle->rowstobuffer = DEFAULT_BUFFERED_ROWS;
    handle->rowheight = 0;
    handle->freezetop = 0;
    handle->sheetopen = 0;
    handle->rowopen = 0;
#ifndef NO_ROW_NUMBERS
    handle->rownr = 0;
#ifndef NO_COLUMN_NUMBERS
    handle->colnr = 0;
#endif
#endif
    //remove filename first if it already exists
    unlink(filename);
    //initialize zip file object
#ifdef USE_MINIZIP
    if ((handle->zip = zipOpen(handle->filename, 0)) == NULL) {
#else
    if ((handle->zip = zip_open(handle->filename, ZIP_CREATE, NULL)) == NULL) {
#endif
      fprintf(stderr, "Error writing to file %s\n", filename);/////
      //unlink(filename);
      free(handle->filename);
      free(handle);
      return NULL;
    }
    //create pipe
    if (pipe(pipefd) != 0) {
      fprintf(stderr, "Error creating pipe\n");/////
      free(handle);
      return NULL;
    }
    handle->pipe_read = fdopen(pipefd[0], "rb");
    handle->pipe_write = fdopen(pipefd[1], "wb");
    //create and start thread that will receive data via pipe
#ifdef USE_WINTHREADS
    if ((handle->thread = CreateThread(NULL, 0, thread_proc, handle, 0, NULL)) == NULL) {
#else
    if (pthread_create(&handle->thread, NULL, thread_proc, handle) != 0) {
#endif
      fprintf(stderr, "Error creating thread\n");/////
#ifdef USE_MINIZIP
      zipClose(handle->zip, NULL);
#else
      zip_close(handle->zip);
#endif
      //unlink(filename);
      free(handle->filename);
      fclose(handle->pipe_read);
      fclose(handle->pipe_write);
      free(handle);
      return NULL;
    }
    //write initial worksheet data
    fprintf(handle->pipe_write, "%s", worksheet_xml_begin);
  }
  return handle;
}

void flush_buffer (xlsxiowriter handle);

DLL_EXPORT_XLSXIO int xlsxiowrite_close (xlsxiowriter handle)
{
  struct column_info_struct* colinfo;
  struct column_info_struct* colinfonext;
  if (!handle)
    return -1;
  //finalize data
  if (handle->pipe_write) {
    //check if buffer should be flushed
    if (!handle->sheetopen)
      flush_buffer(handle);
    //close row if needed
    if (handle->rowopen)
      fprintf(handle->pipe_write, "</row>" OPTIONAL_LINE_BREAK);
    //write worksheet data
    fprintf(handle->pipe_write, "%s", worksheet_xml_end);
    //close pipe
    fclose(handle->pipe_write);
  }
  //wait for thread to finish
#ifdef USE_WINTHREADS
  WaitForSingleObject(handle->thread, INFINITE);
#else
  pthread_join(handle->thread, NULL);
#endif
  //clean up
  colinfo = handle->columninfo;
  while (colinfo) {
    colinfonext = colinfo->next;
    free(colinfo);
    colinfo = colinfonext;
  }
  free(handle->filename);
  free(handle->sheetname);
  if (handle->zip)
#ifdef USE_MINIZIP
    zipClose(handle->zip, NULL);
#else
    zip_close(handle->zip);
#endif
  if (handle->pipe_read)
    fclose(handle->pipe_read);
  free(handle);
  return 0;
}

#ifndef WITHOUT_XLSX_STYLES
#define STYLE_ATTR_HELPER(x) #x
#define STYLE_ATTR(style) " s=\"" STYLE_ATTR_HELPER(style) "\""
#else
#define STYLE_ATTR(style) ""
#endif

//output start of row
void write_row_start (xlsxiowriter handle, const char* rowattr)
{
#ifndef NO_ROW_NUMBERS
  handle->rownr++;
#ifndef NO_COLUMN_NUMBERS
  handle->colnr = 0;
#endif
#endif
  if (handle->sheetopen) {
    if (!handle->rowheight)
      fprintf(handle->pipe_write, "<row%s" ROWNRTAG ">", (rowattr ? rowattr : "") ROWNRPARAM(handle));
    else
      fprintf(handle->pipe_write, "<row ht=\"%.6G\" customHeight=\"1\"%s" ROWNRTAG ">", CALCULATE_COLUMN_HEIGHT(handle->rowheight), (rowattr ? rowattr : "") ROWNRPARAM(handle));
  } else {
    if (!handle->rowheight)
      append_data(&handle->buf, &handle->buflen, "<row%s" ROWNRTAG ">", (rowattr ? rowattr : "") ROWNRPARAM(handle));
    else
      append_data(&handle->buf, &handle->buflen, "<row ht=\"%.6G\" customHeight=\"1\"%s" ROWNRTAG ">",  CALCULATE_COLUMN_HEIGHT(handle->rowheight), (rowattr ? rowattr : "") ROWNRPARAM(handle));
  }
  handle->rowopen = 1;
}

//output cell data
void write_cell_data (xlsxiowriter handle, const char* rowattr, const char* prefix, const char* suffix, const char* format, ...)
{
  va_list args;
#if !defined(NO_ROW_NUMBERS) && !defined(NO_COLUMN_NUMBERS)
  char* cellcoord;
#endif
  if (!handle)
    return;
  //start new row if needed
  if (!handle->rowopen)
    write_row_start(handle, rowattr);
  //get formatted data
  int datalen;
  char* data;
  va_start(args, format);
  if (format && (datalen = vsnprintf(NULL, 0, format, args)) >= 0 && (data = (char*)malloc(datalen + 1)) != NULL) {
    va_end(args);
    va_start(args, format);
    vsnprintf(data, datalen + 1, format, args);
    //prepare data for XML output
    fix_xml_special_chars(&data);
  } else {
    data = NULL;
    datalen = 0;
  }
  va_end(args);
  //determine cell coordinate
#if !defined(NO_ROW_NUMBERS) && !defined(NO_COLUMN_NUMBERS)
  cellcoord = get_A1col(++handle->colnr);
#define COLNRPARAM(handle) , cellcoord, handle->rownr
#else
#define COLNRPARAM(handle)
#endif
  //add cell data
  if (handle->sheetopen) {
    //write cell data
    if (prefix)
      fprintf(handle->pipe_write, prefix COLNRPARAM(handle));
    if (data)
      fprintf(handle->pipe_write, "%s", data);
    if (suffix)
      fprintf(handle->pipe_write, "%s", suffix);
  } else {
    //add cell data to buffer
    if (prefix)
      append_data(&handle->buf, &handle->buflen, prefix COLNRPARAM(handle));
    if (data)
      append_data(&handle->buf, &handle->buflen, "%s", data);
    if (suffix)
      append_data(&handle->buf, &handle->buflen, suffix);
    //collect cell information
    if (!handle->sheetopen) {
      if (!*handle->pcurrentcolumn) {
        //create new column information structure
        struct column_info_struct* colinfo;
        if ((colinfo = (struct column_info_struct*)malloc(sizeof(struct column_info_struct))) != NULL) {
          colinfo->width = 0;
          colinfo->maxwidth = 0;
          colinfo->next = NULL;
          *handle->pcurrentcolumn = colinfo;
        }
      }
      //keep track of biggest column width
      if (data) {
        //only count first line in multiline data
        char* p = strchr(data, '\n');
        if (p)
          datalen = p - data;
        //remember this length if it is the longest one so far
        if (datalen > 0 && datalen > (*handle->pcurrentcolumn)->maxwidth)
          (*handle->pcurrentcolumn)->maxwidth = datalen;
      }
      //prepare for the next column
      handle->pcurrentcolumn = &(*handle->pcurrentcolumn)->next;
    }
  }
#if !defined(NO_ROW_NUMBERS) && !defined(NO_COLUMN_NUMBERS)
  free(cellcoord);
#endif
  free(data);
}

//output buffered data and stop buffering
void flush_buffer (xlsxiowriter handle)
{
  //write section to freeze top row
  if (handle->freezetop > 0)
    fprintf(handle->pipe_write, "%s", worksheet_xml_freeze_top_row);
  //default to row height of 1 line
  //fprintf(handle->pipe_write, "<sheetFormatPr defaultRowHeight=\"%.6G\" customHeight=\"1\"/> OPTIONAL_LINE_BREAK", (double)12.75);
  //write column information
  if (handle->columninfo) {
    int col = 0;
    int len;
    struct column_info_struct* colinfo = handle->columninfo;
    fprintf(handle->pipe_write, "<cols>");
    while (colinfo) {
      ++col;
      //determine column width
      len = colinfo->width;
      if (len == 0) {
        //use detected maximum length if column width specified was zero
        if (colinfo->maxwidth > 0)
          len = colinfo->maxwidth;
      } else if (len < 0) {
        //use detected maximum length if column width specified was negative and the detected maximum length is larger than the absolute value of the specified width
        len = -len;
        if (colinfo->maxwidth > len)
          len = colinfo->maxwidth;
      }
      if (len)
        fprintf(handle->pipe_write, "<col min=\"%i\" max=\"%i\" width=\"%.6G\" customWidth=\"1\"/>" OPTIONAL_LINE_BREAK, col, col, CALCULATE_COLUMN_WIDTH(len));
      else
        fprintf(handle->pipe_write, "<col min=\"%i\" max=\"%i\"/>" OPTIONAL_LINE_BREAK, col, col);
      colinfo = colinfo->next;
    }
    fprintf(handle->pipe_write, "</cols>" OPTIONAL_LINE_BREAK);
  }
  //write initial data
  fprintf(handle->pipe_write, "%s", worksheet_xml_start_data);
  //write buffer and clear it
  if (handle->buf) {
    if (handle->buflen > 0)
      fwrite(handle->buf, 1, handle->buflen, handle->pipe_write);
    free(handle->buf);
    handle->buf = NULL;
  }
  handle->buflen = 0;
  handle->sheetopen = 1;
}

DLL_EXPORT_XLSXIO void xlsxiowrite_set_detection_rows (xlsxiowriter handle, size_t rows)
{
  //abort if currently not buffering
  if (!handle->rowstobuffer || handle->sheetopen)
    return;
  //set number of rows to buffer
  handle->rowstobuffer = rows;
  //flush when zero was specified
  if (!rows)
    flush_buffer(handle);
}

DLL_EXPORT_XLSXIO void xlsxiowrite_set_row_height (xlsxiowriter handle, size_t height)
{
  handle->rowheight = height;
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_column (xlsxiowriter handle, const char* value, int width)
{
  struct column_info_struct** pcolinfo = handle->pcurrentcolumn;
  if (value) {
    if (need_space_preserve_attr(value))
      write_cell_data(handle, STYLE_ATTR(STYLE_HEADER), "<c t=\"inlineStr\"" STYLE_ATTR(STYLE_HEADER) COLNRTAG "><is xml:space=\"preserve\"><t>", "</t></is></c>", "%s", value);
    else
      write_cell_data(handle, STYLE_ATTR(STYLE_HEADER), "<c t=\"inlineStr\"" STYLE_ATTR(STYLE_HEADER) COLNRTAG "><is><t>", "</t></is></c>", "%s", value);
  } else {
    write_cell_data(handle, STYLE_ATTR(STYLE_HEADER), "<c" STYLE_ATTR(STYLE_HEADER) COLNRTAG "/>", NULL, NULL);
  }
  if (*pcolinfo)
    (*pcolinfo)->width = width;
  if (handle->freezetop == 0)
    handle->freezetop = 1;
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_string (xlsxiowriter handle, const char* value)
{
  if (value) {
    if (need_space_preserve_attr(value))
      write_cell_data(handle, NULL, "<c t=\"inlineStr\"" STYLE_ATTR(STYLE_TEXT) COLNRTAG "><is xml:space=\"preserve\"><t>", "</t></is></c>", "%s", value);
    else
      write_cell_data(handle, NULL, "<c t=\"inlineStr\"" STYLE_ATTR(STYLE_TEXT) COLNRTAG "><is><t>", "</t></is></c>", "%s", value);
  } else {
    write_cell_data(handle, NULL, "<c" STYLE_ATTR(STYLE_TEXT) COLNRTAG "/>", NULL, NULL);
  }
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_int (xlsxiowriter handle, int64_t value)
{
  write_cell_data(handle, NULL, "<c" STYLE_ATTR(STYLE_INTEGER) COLNRTAG "><v>", "</v></c>", "%" PRIi64, value);
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_float (xlsxiowriter handle, double value)
{
  write_cell_data(handle, NULL, "<c" STYLE_ATTR(STYLE_GENERAL) COLNRTAG "><v>", "</v></c>", "%.32G", value);
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_datetime (xlsxiowriter handle, time_t value)
{
  double timestamp = ((double)(value) + .499) / 86400 + 25569; //conversion from Unix to Excel timestamp
  write_cell_data(handle, NULL, "<c" STYLE_ATTR(STYLE_DATETIME) COLNRTAG "><v>", "</v></c>", "%.16G", timestamp);
}
/*
Windows (And Mac Office 2011+):

    Unix Timestamp = (Excel Timestamp - 25569) * 86400
    Excel Timestamp = (Unix Timestamp / 86400) + 25569

MAC OS X (pre Office 2011):

    Unix Timestamp = (Excel Timestamp - 24107) * 86400
    Excel Timestamp = (Unix Timestamp / 86400) + 24107
*/

DLL_EXPORT_XLSXIO void xlsxiowrite_next_row (xlsxiowriter handle)
{
  if (!handle)
    return;
  //check if buffer should be flushed
  if (!handle->sheetopen) {
    if (handle->rowstobuffer > 0) {
      if (--handle->rowstobuffer == 0) {
        flush_buffer(handle);
      } else {
      }
    }
  }
  //start new row if needed
  if (!handle->rowopen)
    write_row_start(handle, NULL);
  //end row
  if (handle->rowstobuffer == 0)
    fprintf(handle->pipe_write, "</row>" OPTIONAL_LINE_BREAK);
  else
    append_data(&handle->buf, &handle->buflen, "</row>" OPTIONAL_LINE_BREAK);
  handle->rowopen = 0;
  handle->pcurrentcolumn = &handle->columninfo;
}

