#include "xlsxio_private.h"
#include "xlsxio_read_sharedstrings.h"
#include "xlsxio_read.h"
#include "xlsxio_version.h"
#include <stdlib.h>
#include <stdio.h>
#include <inttypes.h>
#include <string.h>
#include <expat.h>

#ifdef USE_MINIZIP
#  include <minizip/unzip.h>
#  define ZIPFILETYPE unzFile
#  define ZIPFILEENTRYTYPE unzFile
#  if defined(_MSC_VER)
#    include <io.h>
#    define IOSIZETYPE int
#    define IOFN(fn) _##fn
#  else
#    include <unistd.h>
#    define IOSIZETYPE ssize_t
#    define IOFN(fn) fn
#  endif
/*
#  if !defined(Z_DEFLATED) && defined(MZ_COMPRESS_METHOD_DEFLATE) // support minizip2 which defines MZ_COMPRESS_METHOD_DEFLATE instead of Z_DEFLATED
#    ifndef ZCALLBACK
#      define ZCALLBACK
#    endif
#    define voidpf void*
#    define uLong  unsigned long
#  endif
*/
#else
#  if (defined(STATIC) || defined(BUILD_XLSXIO_STATIC) || defined(BUILD_XLSXIO_STATIC_DLL) || (defined(BUILD_XLSXIO) && !defined(BUILD_XLSXIO_DLL) && !defined(BUILD_XLSXIO_SHARED))) && !defined(ZIP_STATIC)
#    define ZIP_STATIC
#  endif
#  include <zip.h>
#  define ZIPFILETYPE zip_t
#  define ZIPFILEENTRYTYPE zip_file_t
#  ifndef USE_LIBZIP
#    define USE_LIBZIP
#  endif
#endif

#if defined(_MSC_VER)
#  undef DLL_EXPORT_XLSXIO
#  define DLL_EXPORT_XLSXIO
#endif

#define PARSE_BUFFER_SIZE 256
//#define PARSE_BUFFER_SIZE 4

static const XLSXIOCHAR* xlsx_content_type = X("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
static const XLSXIOCHAR* xlsm_content_type = X("application/vnd.ms-excel.sheet.macroEnabled.main+xml");
static const XLSXIOCHAR* xltx_content_type = X("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml");
static const XLSXIOCHAR* xltm_content_type = X("application/vnd.ms-excel.template.macroEnabled.main+xml");

#if !defined(XML_UNICODE_WCHAR_T) && !defined(XML_UNICODE)

//UTF-8 version
#define XML_Char_dupchar strdup

static ZIPFILEENTRYTYPE* XML_Char_openzip (ZIPFILETYPE* archive, const XML_Char* filename, int flags)
{
  if (!filename || !*filename)
    return NULL;
#ifdef USE_MINIZIP
  if (unzLocateFile(archive, filename, 0) != UNZ_OK)
    return NULL;
  if (unzOpenCurrentFile(archive) != UNZ_OK)
    return NULL;
  return archive;
#else
  return zip_fopen(archive, filename, flags);
#endif
}

#else

//UTF-16 version
static XML_Char* XML_Char_dupchar(const char* s)
{
  size_t len;
  XML_Char* result;
  if (!s || (len = mbstowcs(NULL, s, 0)) < 0)
    return NULL;
  if ((result = XML_Char_malloc(len + 1)) != NULL) {
    if ((mbstowcs(result, s, len + 1) != len)) {
      free(result);
      return NULL;
    }
  }
  return result;
}

static char* chardupXML_Char(const XML_Char* s)
{
  size_t len;
  char* result;
  if (!s || (len = wcstombs(NULL, s, 0)) == -1)
    return NULL;
  if ((result = (char*)malloc(len + 1)) != NULL) {
    if ((wcstombs(result, s, len + 1) != len)) {
      free(result);
      return NULL;
    }
  }
  return result;
}

static ZIPFILEENTRYTYPE* XML_Char_openzip (ZIPFILETYPE* archive, const XML_Char* filename, int flags)
{
  ZIPFILEENTRYTYPE* result;
  char* s;
  if (!filename || !*filename)
    return NULL;
  if ((s = chardupXML_Char(filename)) == NULL)
    return NULL;
#ifdef USE_MINIZIP
  if (unzLocateFile(archive, s, 0) != UNZ_OK)
    result = NULL;
  else if (unzOpenCurrentFile(archive) != UNZ_OK)
    result = NULL;
  else
    result = archive;
#else
  result = zip_fopen(archive, s, flags);
#endif
  free(s);
  return result;
}

#endif


DLL_EXPORT_XLSXIO void xlsxioread_get_version (int* pmajor, int* pminor, int* pmicro)
{
  if (pmajor)
    *pmajor = XLSXIO_VERSION_MAJOR;
  if (pminor)
    *pminor = XLSXIO_VERSION_MINOR;
  if (pmicro)
    *pmicro = XLSXIO_VERSION_MICRO;
}

DLL_EXPORT_XLSXIO const XLSXIOCHAR* xlsxioread_get_version_string ()
{
  return (const XLSXIOCHAR*)XLSXIO_VERSION_STRING;
}

////////////////////////////////////////////////////////////////////////

//process XML file contents
int expat_process_zip_file (ZIPFILETYPE* zip, const XML_Char* filename, XML_StartElementHandler start_handler, XML_EndElementHandler end_handler, XML_CharacterDataHandler data_handler, void* callbackdata, XML_Parser* xmlparser)
{
  ZIPFILEENTRYTYPE* zipfile;
  XML_Parser parser;
  void* buf;
#ifdef USE_MINIZIP
  int buflen;
#else
  zip_int64_t buflen;
#endif
  int done;
  enum XML_Status status = XML_STATUS_ERROR;
  if ((zipfile = XML_Char_openzip(zip, filename, 0)) == NULL) {
    return -1;
  }
  parser = XML_ParserCreate(NULL);
  XML_SetUserData(parser, callbackdata);
  XML_SetElementHandler(parser, start_handler, end_handler);
  XML_SetCharacterDataHandler(parser, data_handler);
  if (xmlparser)
    *xmlparser = parser;
  buf = XML_GetBuffer(parser, PARSE_BUFFER_SIZE);
#ifdef USE_MINIZIP
    while (buf && (buflen = unzReadCurrentFile(zip, buf, PARSE_BUFFER_SIZE)) >= 0) {
#else
    while (buf && (buflen = zip_fread(zipfile, buf, PARSE_BUFFER_SIZE)) >= 0) {
#endif
      done = buflen < PARSE_BUFFER_SIZE;
      if ((status = XML_ParseBuffer(parser, (int)buflen, (done ? 1 : 0))) == XML_STATUS_ERROR) {
        break;
      }
      if (xmlparser && status == XML_STATUS_SUSPENDED)
        return 0;
      if (done)
        break;
      buf = XML_GetBuffer(parser, PARSE_BUFFER_SIZE);
    }
  XML_ParserFree(parser);
#ifdef USE_MINIZIP
  unzCloseCurrentFile(zip);
#else
  zip_fclose(zipfile);
#endif
  //return (status == XML_STATUS_ERROR != XML_ERROR_FINISHED ? 1 : 0);
  return 0;
}

XML_Parser expat_process_zip_file_suspendable (ZIPFILEENTRYTYPE* zipfile, XML_StartElementHandler start_handler, XML_EndElementHandler end_handler, XML_CharacterDataHandler data_handler, void* callbackdata)
{
  XML_Parser result;
  if ((result = XML_ParserCreate(NULL)) != NULL) {
    XML_SetUserData(result, callbackdata);
    XML_SetElementHandler(result, start_handler, end_handler);
    XML_SetCharacterDataHandler(result, data_handler);
  }
  return result;
}

enum XML_Status expat_process_zip_file_resume (ZIPFILEENTRYTYPE* zipfile, XML_Parser xmlparser)
{
  enum XML_Status status;
  status = XML_ResumeParser(xmlparser);
  if (status == XML_STATUS_SUSPENDED)
    return status;
  if (status == XML_STATUS_ERROR && XML_GetErrorCode(xmlparser) != XML_ERROR_NOT_SUSPENDED)
    return status;
  void* buf;
#ifdef USE_MINIZIP
  int buflen;
#else
  zip_int64_t buflen;
#endif
  int done;
  buf = XML_GetBuffer(xmlparser, PARSE_BUFFER_SIZE);
#ifdef USE_MINIZIP
  while (buf && (buflen = unzReadCurrentFile(zipfile, buf, PARSE_BUFFER_SIZE)) >= 0) {
#else
  while (buf && (buflen = zip_fread(zipfile, buf, PARSE_BUFFER_SIZE)) >= 0) {
#endif
    done = buflen < PARSE_BUFFER_SIZE;
    if ((status = XML_ParseBuffer(xmlparser, (int)buflen, (done ? 1 : 0))) == XML_STATUS_ERROR)
      return status;
    if (status == XML_STATUS_SUSPENDED)
      return status;
    if (done)
      break;
    buf = XML_GetBuffer(xmlparser, PARSE_BUFFER_SIZE);
  }
  //XML_ParserFree(xmlparser);
  return status;
}

//compare XML name ignoring case and ignoring namespace (returns 0 on match)
#ifdef ASSUME_NO_NAMESPACE
#define XML_Char_icmp_ins XML_Char_icmp
#else
int XML_Char_icmp_ins (const XML_Char* value, const XML_Char* name)
{
  size_t valuelen = XML_Char_len(value);
  size_t namelen = XML_Char_len(name);
  if (valuelen == namelen)
    return XML_Char_icmp(value, name);
  if (valuelen > namelen) {
    if (value[valuelen - namelen - 1] != ':')
      return 1;
    return XML_Char_icmp(value + (valuelen - namelen), name);
  }
  return -1;
}
#endif

//get expat attribute by name, returns NULL if not found
const XML_Char* get_expat_attr_by_name (const XML_Char** atts, const XML_Char* name)
{
  const XML_Char** p = atts;
  if (p) {
    while (*p) {
      //if (XML_Char_icmp(*p++, name) == 0)
      if (XML_Char_icmp_ins(*p++, name) == 0)
        return *p;
      p++;
    }
  }
  return NULL;
}

//generate .rels filename, returns NULL on error, caller must free result
XML_Char* get_relationship_filename (const XML_Char* filename)
{
  XML_Char* result;
  size_t filenamelen = XML_Char_len(filename);
  if ((result = XML_Char_malloc(filenamelen + 12)) != NULL) {
    size_t i = filenamelen;
    while (i > 0) {
      if (filename[i - 1] == '/')
        break;
      i--;
    }
    XML_Char_poscpy(result, 0, filename, i);
    XML_Char_poscpy(result, i, X("_rels/"), 6);
    XML_Char_poscpy(result, i + 6, filename + i, filenamelen - i);
    XML_Char_poscpy(result, filenamelen + 6, X(".rels"), 6);
  }
  return result;
}

//join basepath and filename (caller must free result)
XML_Char* join_basepath_filename (const XML_Char* basepath, const XML_Char* filename)
{
  XML_Char* result = NULL;
  if (filename && *filename) {
    if (filename[0] == '/' && filename[1]) {
      //file is absolute: remove leading slash
      result = XML_Char_dup(filename + 1);
    } else {
      //file is relative: prepend base path
      size_t basepathlen = (basepath ? XML_Char_len(basepath) : 0);
      size_t filenamelen = XML_Char_len(filename);
      if ((result = XML_Char_malloc(basepathlen + filenamelen + 1)) != NULL) {
        if (basepathlen > 0)
          XML_Char_poscpy(result, 0, basepath, basepathlen);
        XML_Char_poscpy(result, basepathlen, filename, filenamelen);
        result[basepathlen + filenamelen] = 0;
      }
    }
  }
  return result;
}

//determine column number based on cell coordinate (e.g. "A1"), returns 1-based column number or 0 on error
size_t get_col_nr (const XML_Char* A1col)
{
  const XML_Char* p = A1col;
  size_t result = 0;
  if (p) {
    while (*p) {
      if (*p >= 'A' && *p <= 'Z')
        result = result * 26 + (*p - 'A') + 1;
      else if (*p >= 'a' && *p <= 'z')
        result = result * 26 + (*p - 'a') + 1;
      else if (*p >= '0' && *p <= '9' && p != A1col)
        return result;
      else
        break;
      p++;
    }
  }
  return 0;
}

//determine row number based on cell coordinate (e.g. "A1"), returns 1-based row number or 0 on error
size_t get_row_nr (const XML_Char* A1col)
{
  const XML_Char* p = A1col;
  size_t result = 0;
  if (p) {
    while (*p) {
      if ((*p >= 'A' && *p <= 'Z') || (*p >= 'a' && *p <= 'z'))
        ;
      else if (*p >= '0' && *p <= '9' && p != A1col)
        result = result * 10 + (*p - '0');
      else
        return 0;
      p++;
    }
  }
  return result;
}

////////////////////////////////////////////////////////////////////////

struct xlsxio_read_struct {
  ZIPFILETYPE* zip;
};

DLL_EXPORT_XLSXIO xlsxioreader xlsxioread_open (const char* filename)
{
  xlsxioreader result;
  if ((result = (xlsxioreader)malloc(sizeof(struct xlsxio_read_struct))) != NULL) {
#ifdef USE_MINIZIP
    if ((result->zip = unzOpen(filename)) == NULL) {
#else
    if ((result->zip = zip_open(filename, ZIP_RDONLY, NULL)) == NULL) {
#endif
      free(result);
      return NULL;
    }
  }
  return result;
}

#ifdef USE_MINIZIP
struct minizip_io_filehandle_data {
  int filehandle;
};

voidpf ZCALLBACK minizip_io_filehandle_open_file_fn (voidpf opaque, const char* filename, int mode)
{
  if (!opaque || ((struct minizip_io_filehandle_data*)opaque)->filehandle < 0)
    return NULL;
  return &((struct minizip_io_filehandle_data*)opaque)->filehandle;
}

uLong ZCALLBACK minizip_io_filehandle_read_file_fn (voidpf opaque, voidpf stream, void* buf, uLong size)
{
  IOSIZETYPE len;
  if (!opaque || !stream || !buf || size == 0)
    return 0;
  if ((len = IOFN(read)(*(int*)stream, buf, size)) < 0)
    return 0;
  return len;
}

/*
uLong ZCALLBACK minizip_io_filehandle_write_file_fn (voidpf opaque, voidpf stream, const void* buf, uLong size)
{
  return 0;
}
*/

int ZCALLBACK minizip_io_filehandle_close_file_fn (voidpf opaque, voidpf stream)
{
  if (stream)
    close(*(int*)stream);
  free(opaque);
  return 0;
}

int ZCALLBACK minizip_io_filehandle_testerror_file_fn (voidpf opaque, voidpf stream)
{
  return 0;
}

long ZCALLBACK minizip_io_filehandle_tell_file_fn (voidpf opaque, voidpf stream)
{
  return IOFN(lseek)(*(int*)stream, 0, SEEK_CUR);
}

long ZCALLBACK minizip_io_filehandle_seek_file_fn (voidpf opaque, voidpf stream, uLong offset, int origin)
{
  int whence;
  if (!opaque || !stream)
    return -1;
  switch (origin) {
    case ZLIB_FILEFUNC_SEEK_CUR :
      whence = SEEK_CUR;
      break;
    case ZLIB_FILEFUNC_SEEK_END :
      whence = SEEK_END;
      break;
    case ZLIB_FILEFUNC_SEEK_SET :
      whence = SEEK_SET;
      break;
    default :
      return -1;
  }
  return (IOFN(lseek)(*(int*)stream, offset, whence) >= 0 ? 0 : -1);
}
#endif

DLL_EXPORT_XLSXIO xlsxioreader xlsxioread_open_filehandle (int filehandle)
{
  xlsxioreader result;
  if ((result = (xlsxioreader)malloc(sizeof(struct xlsxio_read_struct))) != NULL) {
#ifdef USE_MINIZIP
    zlib_filefunc_def minizip_io_filehandle_functions;
    if ((minizip_io_filehandle_functions.opaque = malloc(sizeof(struct minizip_io_filehandle_data))) == NULL) {
      free(result);
      return NULL;
    }
    minizip_io_filehandle_functions.zopen_file = minizip_io_filehandle_open_file_fn;
    minizip_io_filehandle_functions.zread_file = minizip_io_filehandle_read_file_fn;
    minizip_io_filehandle_functions.zwrite_file = /*minizip_io_filehandle_write_file_fn*/NULL;
    minizip_io_filehandle_functions.ztell_file = minizip_io_filehandle_tell_file_fn;
    minizip_io_filehandle_functions.zseek_file = minizip_io_filehandle_seek_file_fn;
    minizip_io_filehandle_functions.zclose_file = minizip_io_filehandle_close_file_fn;
    minizip_io_filehandle_functions.zerror_file = minizip_io_filehandle_testerror_file_fn;
    ((struct minizip_io_filehandle_data*)minizip_io_filehandle_functions.opaque)->filehandle = filehandle;
    if ((result->zip = unzOpen2(NULL, &minizip_io_filehandle_functions)) == NULL) {
      free(result);
      return NULL;
    }
#else
    if ((result->zip = zip_fdopen(filehandle, ZIP_RDONLY, NULL)) == NULL) {
      free(result);
      return NULL;
    }
#endif
  }
  return result;
}

#ifdef USE_MINIZIP
struct minizip_io_memory_data {
  void* data;
  uint64_t datalen;
  int freedata;
};

struct minizip_io_memory_handle {
  uint64_t pos;
};

voidpf ZCALLBACK minizip_io_memory_open_file_fn (voidpf opaque, const char* filename, int mode)
{
  struct minizip_io_memory_handle* result;
  if (!opaque || !((struct minizip_io_memory_data*)opaque)->data)
    return NULL;
  if ((result = (struct minizip_io_memory_handle*)malloc(sizeof(struct minizip_io_memory_handle))) != NULL) {
    result->pos = 0;
  }
  return result;
}

uLong ZCALLBACK minizip_io_memory_read_file_fn (voidpf opaque, voidpf stream, void* buf, uLong size)
{
  uLong len;
  if (!opaque || !stream || !buf || size == 0)
    return 0;
  if (((struct minizip_io_memory_handle*)stream)->pos + size <= ((struct minizip_io_memory_data*)opaque)->datalen)
    len = size;
  else
    len = ((struct minizip_io_memory_data*)opaque)->datalen - ((struct minizip_io_memory_handle*)stream)->pos;
  memcpy(buf, (char *)(((struct minizip_io_memory_data*)opaque)->data) + ((struct minizip_io_memory_handle*)stream)->pos, len);
  ((struct minizip_io_memory_handle*)stream)->pos += len;
  return len;
}

/*
uLong ZCALLBACK minizip_io_memory_write_file_fn (voidpf opaque, voidpf stream, const void* buf, uLong size)
{
  return 0;
}
*/

int ZCALLBACK minizip_io_memory_close_file_fn (voidpf opaque, voidpf stream)
{
  free(stream);
  if (opaque && ((struct minizip_io_memory_data*)opaque)->freedata)
    free(((struct minizip_io_memory_data*)opaque)->data);
  free(opaque);
  return 0;
}

int ZCALLBACK minizip_io_memory_testerror_file_fn (voidpf opaque, voidpf stream)
{
  return 0;
}

long ZCALLBACK minizip_io_memory_tell_file_fn (voidpf opaque, voidpf stream)
{
  if (!opaque || !stream)
    return 0;
  return ((struct minizip_io_memory_handle*)stream)->pos;
}

long ZCALLBACK minizip_io_memory_seek_file_fn (voidpf opaque, voidpf stream, uLong offset, int origin)
{
  switch (origin) {
    case ZLIB_FILEFUNC_SEEK_CUR :
      /*if (offset < 0) {
        if (((struct minizip_io_memory_handle*)stream)->pos < -offset)
          ((struct minizip_io_memory_handle*)stream)->pos = 0;
        else
          ((struct minizip_io_memory_handle*)stream)->pos += offset;
      } else*/ {
        if (((struct minizip_io_memory_handle*)stream)->pos + offset > ((struct minizip_io_memory_data*)opaque)->datalen)
          ((struct minizip_io_memory_handle*)stream)->pos = ((struct minizip_io_memory_data*)opaque)->datalen;
        else
          ((struct minizip_io_memory_handle*)stream)->pos += offset;
      }
      break;
    case ZLIB_FILEFUNC_SEEK_END :
      /*if (offset < 0) {
        if (((struct minizip_io_memory_data*)opaque)->datalen < -offset)
          ((struct minizip_io_memory_handle*)stream)->pos = 0;
        else
          ((struct minizip_io_memory_handle*)stream)->pos = ((struct minizip_io_memory_data*)opaque)->datalen + offset;
      } else*/ {
        ((struct minizip_io_memory_handle*)stream)->pos = ((struct minizip_io_memory_data*)opaque)->datalen;
      }
      break;
    case ZLIB_FILEFUNC_SEEK_SET :
      /*if (offset < 0) {
        ((struct minizip_io_memory_handle*)stream)->pos = 0;
      } else*/ {
        if (offset > ((struct minizip_io_memory_data*)opaque)->datalen)
          ((struct minizip_io_memory_handle*)stream)->pos = ((struct minizip_io_memory_data*)opaque)->datalen;
        else
          ((struct minizip_io_memory_handle*)stream)->pos = offset;
      }
      ((struct minizip_io_memory_handle*)stream)->pos = offset;
      break;
    default :
      return -1;
  }
  return 0;
}
#endif

DLL_EXPORT_XLSXIO xlsxioreader xlsxioread_open_memory (void* data, uint64_t datalen, int freedata)
{
  xlsxioreader result;
#ifdef USE_MINIZIP
  if ((result = (xlsxioreader)malloc(sizeof(struct xlsxio_read_struct))) != NULL) {
    zlib_filefunc_def minizip_io_memory_functions;
    if ((minizip_io_memory_functions.opaque = malloc(sizeof(struct minizip_io_memory_data))) == NULL) {
      free(result);
      return NULL;
    }
    minizip_io_memory_functions.zopen_file = minizip_io_memory_open_file_fn;
    minizip_io_memory_functions.zread_file = minizip_io_memory_read_file_fn;
    minizip_io_memory_functions.zwrite_file = /*minizip_io_memory_write_file_fn*/NULL;
    minizip_io_memory_functions.ztell_file = minizip_io_memory_tell_file_fn;
    minizip_io_memory_functions.zseek_file = minizip_io_memory_seek_file_fn;
    minizip_io_memory_functions.zclose_file = minizip_io_memory_close_file_fn;
    minizip_io_memory_functions.zerror_file = minizip_io_memory_testerror_file_fn;
    ((struct minizip_io_memory_data*)minizip_io_memory_functions.opaque)->data = data;
    ((struct minizip_io_memory_data*)minizip_io_memory_functions.opaque)->datalen = datalen;
    ((struct minizip_io_memory_data*)minizip_io_memory_functions.opaque)->freedata = freedata;
    if ((result->zip = unzOpen2(NULL, &minizip_io_memory_functions)) == NULL) {
      free(result);
      return NULL;
    }
  }
#else
  zip_source_t* zipsrc;
  if ((zipsrc = zip_source_buffer_create(data, datalen, freedata, NULL)) == NULL) {
    return NULL;
  }
  if ((result = (xlsxioreader)malloc(sizeof(struct xlsxio_read_struct))) != NULL) {
    if ((result->zip = zip_open_from_source(zipsrc, ZIP_RDONLY, NULL)) == NULL) {
      zip_source_free(zipsrc);
      free(result);
      return NULL;
    }
  }
#endif
  return result;
}

DLL_EXPORT_XLSXIO void xlsxioread_close (xlsxioreader handle)
{
  if (handle) {
    //note: no need to call zip_source_free() after successful use in zip_open_from_source()
#ifdef USE_MINIZIP
    unzClose(handle->zip);
#else
    zip_close(handle->zip);
#endif
    free(handle);
  }
}

////////////////////////////////////////////////////////////////////////

//callback function definition for use with iterate_files_by_contenttype
typedef void (*contenttype_file_callback_fn)(ZIPFILETYPE* zip, const XML_Char* filename, const XML_Char* contenttype, void* callbackdata);

struct iterate_files_by_contenttype_callback_data {
  ZIPFILETYPE* zip;
  const XML_Char* contenttype;
  contenttype_file_callback_fn filecallbackfn;
  void* filecallbackdata;
};

//expat callback function for element start used by iterate_files_by_contenttype
void iterate_files_by_contenttype_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct iterate_files_by_contenttype_callback_data* data = (struct iterate_files_by_contenttype_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("Override")) == 0) {
    //explicitly specified file
    const XML_Char* contenttype;
    const XML_Char* partname;
    if ((contenttype = get_expat_attr_by_name(atts, X("ContentType"))) != NULL && XML_Char_icmp(contenttype, data->contenttype) == 0) {
      if ((partname = get_expat_attr_by_name(atts, X("PartName"))) != NULL) {
        if (partname[0] == '/')
          partname++;
        data->filecallbackfn(data->zip, partname, contenttype, data->filecallbackdata);
      }
    }
  } else if (XML_Char_icmp_ins(name, X("Default")) == 0) {
    //by extension
    const XML_Char* contenttype;
    const XML_Char* extension;
    if ((contenttype = get_expat_attr_by_name(atts, X("ContentType"))) != NULL && XML_Char_icmp(contenttype, data->contenttype) == 0) {
      if ((extension = get_expat_attr_by_name(atts, X("Extension"))) != NULL) {
        XML_Char* filename;
        size_t filenamelen;
        size_t extensionlen = XML_Char_len(extension);
#ifdef USE_MINIZIP
#define UNZIP_FILENAME_BUFFER_STEP 32
        char* buf;
        size_t buflen;
        int status;
unz_global_info zipglobalinfo;
unzGetGlobalInfo(data->zip, &zipglobalinfo);
        buf = (char*)malloc(buflen = UNZIP_FILENAME_BUFFER_STEP);
        status = unzGoToFirstFile(data->zip);
        while (status == UNZ_OK) {
          buf[buflen - 1] = 0;
          while ((status = unzGetCurrentFileInfo(data->zip, NULL, buf, buflen, NULL, 0, NULL, 0)) == UNZ_OK && buf[buflen - 1] != 0) {
            buflen += UNZIP_FILENAME_BUFFER_STEP;
            buf = (char*)realloc(buf, buflen);
            buf[buflen - 1] = 0;
          }
          if (status != UNZ_OK)
            break;
          filename = XML_Char_dupchar(buf);
          status = unzGoToNextFile(data->zip);
#else
        zip_int64_t i;
        zip_int64_t zipnumfiles = zip_get_num_entries(data->zip, 0);
        for (i = 0; i < zipnumfiles; i++) {
          filename = XML_Char_dupchar(zip_get_name(data->zip, i, ZIP_FL_ENC_GUESS));
#endif
          filenamelen = XML_Char_len(filename);
          if (filenamelen > extensionlen && filename[filenamelen - extensionlen - 1] == '.' && XML_Char_icmp(filename + filenamelen - extensionlen, extension) == 0) {
            data->filecallbackfn(data->zip, filename, contenttype, data->filecallbackdata);
          }
          free(filename);
        }
#ifdef USE_MINIZIP
        free(buf);
#endif
      }
    }
  }
}

//list file names by content type
int iterate_files_by_contenttype (ZIPFILETYPE* zip, const XML_Char* contenttype, contenttype_file_callback_fn filecallbackfn, void* filecallbackdata, XML_Parser* xmlparser)
{
  struct iterate_files_by_contenttype_callback_data callbackdata = {
    .zip = zip,
    .contenttype = contenttype,
    .filecallbackfn = filecallbackfn,
    .filecallbackdata = filecallbackdata
  };
  return expat_process_zip_file(zip, X("[Content_Types].xml"), iterate_files_by_contenttype_expat_callback_element_start, NULL, NULL, &callbackdata, xmlparser);
}

////////////////////////////////////////////////////////////////////////

//callback structure used by main_sheet_list_expat_callback_element_start
struct main_sheet_list_callback_data {
  XML_Parser xmlparser;
  xlsxioread_list_sheets_callback_fn callback;
  void* callbackdata;
};

//callback used by xlsxioread_list_sheets
void main_sheet_list_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_list_callback_data* data = (struct main_sheet_list_callback_data*)callbackdata;
  if (data && data->callback) {
    if (XML_Char_icmp_ins(name, X("sheet")) == 0) {
      const XML_Char* sheetname;
      //const XML_Char* relid = get_expat_attr_by_name(atts, X("r:id"));
      if ((sheetname = get_expat_attr_by_name(atts, X("name"))) != NULL) {
        if (data->callback) {
          if ((*data->callback)(sheetname, data->callbackdata) != 0) {
            XML_StopParser(data->xmlparser, XML_FALSE);
            return;
          }
/*
        } else {
          //for non-calback method suspend here
          XML_StopParser(data->xmlparser, XML_TRUE);
*/
        }
      }
    }
  }
}

//process contents each sheet listed in main sheet
void xlsxioread_list_sheets_callback (ZIPFILETYPE* zip, const XML_Char* filename, const XML_Char* contenttype, void* callbackdata)
{
  //get sheet information from file
  expat_process_zip_file(zip, filename, main_sheet_list_expat_callback_element_start, NULL, NULL, callbackdata, &((struct main_sheet_list_callback_data*)callbackdata)->xmlparser);
}

//list all worksheets
DLL_EXPORT_XLSXIO void xlsxioread_list_sheets (xlsxioreader handle, xlsxioread_list_sheets_callback_fn callback, void* callbackdata)
{
  if (!handle || !callback)
    return;
  //process contents of main sheet
  struct main_sheet_list_callback_data sheetcallbackdata = {
    .xmlparser = NULL,
    .callback = callback,
    .callbackdata = callbackdata
  };
  iterate_files_by_contenttype(handle->zip, xlsx_content_type, xlsxioread_list_sheets_callback, &sheetcallbackdata, &sheetcallbackdata.xmlparser);
  iterate_files_by_contenttype(handle->zip, xlsm_content_type, xlsxioread_list_sheets_callback, &sheetcallbackdata, &sheetcallbackdata.xmlparser);
  iterate_files_by_contenttype(handle->zip, xltx_content_type, xlsxioread_list_sheets_callback, &sheetcallbackdata, &sheetcallbackdata.xmlparser);
  iterate_files_by_contenttype(handle->zip, xltm_content_type, xlsxioread_list_sheets_callback, &sheetcallbackdata, &sheetcallbackdata.xmlparser);
}

////////////////////////////////////////////////////////////////////////

//callback data structure used by main_sheet_get_sheetfile_callback
struct main_sheet_get_rels_callback_data {
  XML_Parser xmlparser;
  const XML_Char* sheetname;
  XML_Char* basepath;
  XML_Char* sheetrelid;
  XML_Char* sheetfile;
  XML_Char* sharedstringsfile;
  XML_Char* stylesfile;
};

//determine relationship id for specific sheet name
void main_sheet_get_relid_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("sheet")) == 0) {
    const XML_Char* sheetname = get_expat_attr_by_name(atts, X("name"));
    if (!data->sheetname || XML_Char_icmp(sheetname, data->sheetname) == 0) {
      const XML_Char* relid = get_expat_attr_by_name(atts, X("r:id"));
      if (relid && *relid) {
        data->sheetrelid = XML_Char_dup(relid);
        XML_StopParser(data->xmlparser, XML_FALSE);
        return;
      }
    }
  }
}

//determine file names for specific relationship id
void main_sheet_get_sheetfile_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (data->sheetrelid) {
    if (XML_Char_icmp_ins(name, X("Relationship")) == 0) {
      const XML_Char* reltype;
      if ((reltype = get_expat_attr_by_name(atts, X("Type"))) != NULL) {
        if (XML_Char_icmp(reltype, X("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")) == 0) {
          const XML_Char* relid = get_expat_attr_by_name(atts, X("Id"));
          if (XML_Char_icmp(relid, data->sheetrelid) == 0) {
            const XML_Char* filename = get_expat_attr_by_name(atts, X("Target"));
            if (filename && *filename) {
              data->sheetfile = join_basepath_filename(data->basepath, filename);
            }
          }
        } else if (XML_Char_icmp(reltype, X("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")) == 0) {
          const XML_Char* filename = get_expat_attr_by_name(atts, X("Target"));
          if (filename && *filename) {
            data->sharedstringsfile = join_basepath_filename(data->basepath, filename);
          }
        } else if (XML_Char_icmp(reltype, X("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")) == 0) {
          const XML_Char* filename = get_expat_attr_by_name(atts, X("Target"));
          if (filename && *filename) {
            data->stylesfile = join_basepath_filename(data->basepath, filename);
          }
        }
      }
    }
  }
}

//determine the file name for a specified sheet name
void main_sheet_get_sheetfile_callback (ZIPFILETYPE* zip, const XML_Char* filename, const XML_Char* contenttype, void* callbackdata)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (!data->sheetrelid) {
    expat_process_zip_file(zip, filename, main_sheet_get_relid_expat_callback_element_start, NULL, NULL, callbackdata, &data->xmlparser);
  }
  if (data->sheetrelid) {
    XML_Char* relfilename;
    //determine base name (including trailing slash)
    size_t i = XML_Char_len(filename);
    while (i > 0) {
      if (filename[i - 1] == '/')
        break;
      i--;
    }
    if (data->basepath)
      free(data->basepath);
    if ((data->basepath = XML_Char_malloc(i + 1)) != NULL) {
      XML_Char_poscpy(data->basepath, 0, filename, i);
      data->basepath[i] = 0;
    }
    //find sheet filename in relationship contents
    if ((relfilename = get_relationship_filename(filename)) != NULL) {
      expat_process_zip_file(zip, relfilename, main_sheet_get_sheetfile_expat_callback_element_start, NULL, NULL, callbackdata, &data->xmlparser);
      free(relfilename);
    } else {
      free(data->sheetrelid);
      data->sheetrelid = NULL;
      if (data->basepath) {
        free(data->basepath);
        data->basepath = NULL;
      }
    }
  }
}

////////////////////////////////////////////////////////////////////////

typedef enum {
  none,
  value_string,
  inline_string,
  shared_string
} cell_string_type_enum;

#define XLSXIOREAD_NO_CALLBACK          0x80

struct data_sheet_callback_data {
  XML_Parser xmlparser;
  struct sharedstringlist* sharedstrings;
  size_t rownr;
  size_t colnr;
  size_t cols;
  XML_Char* celldata;
  size_t celldatalen;
  cell_string_type_enum cell_string_type;
  unsigned int flags;
  XML_Char* skiptag;                        //tag to skip
  size_t skiptagcount;                  //nesting level for current tag to skip
  XML_StartElementHandler skip_start;   //start handler to set after skipping
  XML_EndElementHandler skip_end;       //end handler to set after skipping
  XML_CharacterDataHandler skip_data;   //data handler to set after skipping
  xlsxioread_process_row_callback_fn sheet_row_callback;
  xlsxioread_process_cell_callback_fn sheet_cell_callback;
  void* callbackdata;
};

void data_sheet_callback_data_initialize (struct data_sheet_callback_data* data, struct sharedstringlist* sharedstrings, unsigned int flags, xlsxioread_process_cell_callback_fn cell_callback, xlsxioread_process_row_callback_fn row_callback, void* callbackdata)
{
  data->xmlparser = NULL;
  data->sharedstrings = sharedstrings;
  data->rownr = 0;
  data->colnr = 0;
  data->cols = 0;
  data->celldata = NULL;
  data->celldatalen = 0;
  data->cell_string_type = none;
  data->flags = flags;
  data->skiptag = NULL;
  data->skiptagcount = 0;
  data->skip_start = NULL;
  data->skip_end = NULL;
  data->skip_data = NULL;
  data->sheet_cell_callback = cell_callback;
  data->sheet_row_callback = row_callback;
  data->callbackdata = callbackdata;
}

void data_sheet_callback_data_cleanup (struct data_sheet_callback_data* data)
{
  sharedstringlist_destroy(data->sharedstrings);
  free(data->celldata);
  free(data->skiptag);
}

void data_sheet_expat_callback_skip_tag_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (name && XML_Char_icmp_ins(name, data->skiptag) == 0) {
    //increment nesting level
    data->skiptagcount++;
  }
}

void data_sheet_expat_callback_skip_tag_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (!name || XML_Char_icmp_ins(name, data->skiptag) == 0) {
    if (--data->skiptagcount == 0) {
      //restore handlers when done skipping
      XML_SetElementHandler(data->xmlparser, data->skip_start, data->skip_end);
      XML_SetCharacterDataHandler(data->xmlparser, data->skip_data);
      free(data->skiptag);
      data->skiptag = NULL;
    }
  }
}

void data_sheet_expat_callback_find_worksheet_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_worksheet_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_sheetdata_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_sheetdata_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_row_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_row_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_cell_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_cell_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_value_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_value_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_value_data (void* callbackdata, const XML_Char* buf, int buflen);
void data_sheet_expat_callback_find_worksheet_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("worksheet")) == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_sheetdata_start, NULL);
  }
}

void data_sheet_expat_callback_find_worksheet_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("worksheet")) == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_worksheet_start, NULL);
  }
}

void data_sheet_expat_callback_find_sheetdata_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("sheetData")) == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_row_start, data_sheet_expat_callback_find_sheetdata_end);
  }
}

void data_sheet_expat_callback_find_sheetdata_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("sheetData")) == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_sheetdata_start, data_sheet_expat_callback_find_worksheet_end);
  } else {
    data_sheet_expat_callback_find_worksheet_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_row_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("row")) == 0) {
    const XML_Char* hidden = get_expat_attr_by_name(atts, X("hidden"));
    if (!hidden || XML_Char_tol(hidden) == 0 || !(data->flags & XLSXIOREAD_SKIP_HIDDEN_ROWS)) {
      //nesting level for current tag to skip
      //start handler to set after skipping
      //end handler to set after skipping
      //data handler to set after skipping

      data->rownr++;
      data->colnr = 0;
      XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_cell_start, data_sheet_expat_callback_find_row_end);
      //for non-calback method suspend here on new row
      if (data->flags & XLSXIOREAD_NO_CALLBACK) {
        XML_StopParser(data->xmlparser, XML_TRUE);
      }
    } else {
      //skip hidden tow
      XML_SetElementHandler(data->xmlparser, NULL, data_sheet_expat_callback_find_row_end);
    }
  }
}

void data_sheet_expat_callback_find_row_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("row")) == 0) {
    //determine number of columns based on first row
    if (data->rownr == 1 && data->cols == 0)
      data->cols = data->colnr;
    //add empty columns if needed
    if (!(data->flags & XLSXIOREAD_NO_CALLBACK) && data->sheet_cell_callback && !(data->flags & XLSXIOREAD_SKIP_EMPTY_CELLS)) {
      while (data->colnr < data->cols) {
        if ((*data->sheet_cell_callback)(data->rownr, data->colnr + 1, NULL, data->callbackdata)) {
          XML_StopParser(data->xmlparser, XML_FALSE);
          return;
        }
        data->colnr++;
      }
    }
    free(data->celldata);
    data->celldata = NULL;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_row_start, data_sheet_expat_callback_find_sheetdata_end);
    //process end of row
    if (!(data->flags & XLSXIOREAD_NO_CALLBACK)) {
      if (data->sheet_row_callback) {
        if ((*data->sheet_row_callback)(data->rownr, data->colnr, data->callbackdata)) {
          XML_StopParser(data->xmlparser, XML_FALSE);
          return;
        }
      }
    } else {
      //for non-calback method suspend here on end of row
      if (data->flags & XLSXIOREAD_NO_CALLBACK) {
        XML_StopParser(data->xmlparser, XML_TRUE);
      }
    }
  } else {
    data_sheet_expat_callback_find_sheetdata_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_cell_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("c")) == 0) {
    const XML_Char* t = get_expat_attr_by_name(atts, X("r"));
    size_t cellcolnr = get_col_nr(t);
    //skip everything when out of bounds
    if (cellcolnr && data->cols && (data->flags & XLSXIOREAD_SKIP_EXTRA_CELLS) && cellcolnr > data->cols) {
      data->colnr = cellcolnr - 1;
      return;
    }
    //insert empty rows if needed
    if (data->colnr == 0) {
      size_t cellrownr = get_row_nr(t);
      if (cellrownr) {
        if (!(data->flags & XLSXIOREAD_SKIP_EMPTY_ROWS) && !(data->flags & XLSXIOREAD_NO_CALLBACK)) {
          while (data->rownr < cellrownr) {
            //insert empty columns
            if (!(data->flags & XLSXIOREAD_SKIP_EMPTY_CELLS) && data->sheet_cell_callback) {
              while (data->colnr < data->cols) {
                if ((*data->sheet_cell_callback)(data->rownr, data->colnr + 1, NULL, data->callbackdata)) {
                  XML_StopParser(data->xmlparser, XML_FALSE);
                  return;
                }
                data->colnr++;
              }
            }
            //finish empty row
            if (data->sheet_row_callback) {
              if ((*data->sheet_row_callback)(data->rownr, data->cols, data->callbackdata)) {
                XML_StopParser(data->xmlparser, XML_FALSE);
                return;
              }
            }
            data->rownr++;
            data->colnr = 0;
          }
        } else {
          data->rownr = cellrownr;
        }
      }
    }
    //insert empty columns if needed
    if (cellcolnr) {
      cellcolnr--;
      if (data->flags & XLSXIOREAD_SKIP_EMPTY_CELLS || data->flags & XLSXIOREAD_NO_CALLBACK) {
        data->colnr = cellcolnr;
      } else {
        while (data->colnr < cellcolnr) {
          if (data->sheet_cell_callback) {
            if ((*data->sheet_cell_callback)(data->rownr, data->colnr + 1, NULL, data->callbackdata)) {
              XML_StopParser(data->xmlparser, XML_FALSE);
              return;
            }
          }
          data->colnr++;
        }
      }
    }
    //determing value type
    if ((t = get_expat_attr_by_name(atts, X("t"))) != NULL && XML_Char_icmp(t, X("s")) == 0)
      data->cell_string_type = shared_string;
    else
      data->cell_string_type = value_string;
    //prepare empty value data
    free(data->celldata);
    data->celldata = NULL;
    data->celldatalen = 0;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_value_start, data_sheet_expat_callback_find_cell_end);
  }
}

void data_sheet_expat_callback_find_cell_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("c")) == 0) {
    //determine value
    if (data->celldata) {
      const XML_Char* s = NULL;
      data->celldata[data->celldatalen] = 0;
      if (data->cell_string_type == shared_string) {
        //get shared string
        XML_Char* p = NULL;
        long num = XML_Char_strtol(data->celldata, &p, 10);
        if (!p || (p != data->celldata && *p == 0)) {
          s = sharedstringlist_get(data->sharedstrings, num);
          free(data->celldata);
          data->celldata = (s ? XML_Char_dup(s) : NULL);
        }
      } else if (data->cell_string_type == none) {
        //unknown value type
        free(data->celldata);
        data->celldata = NULL;
      }
    }
    //reset data
    data->colnr++;
    data->cell_string_type = none;
    data->celldatalen = 0;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_cell_start, data_sheet_expat_callback_find_row_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
    //process data if needed
    if (!(data->cols && (data->flags & XLSXIOREAD_SKIP_EXTRA_CELLS) && data->colnr > data->cols)) {
      //process data
      if (!(data->flags & XLSXIOREAD_NO_CALLBACK)) {
        if (data->sheet_cell_callback) {
          if ((*data->sheet_cell_callback)(data->rownr, data->colnr, data->celldata, data->callbackdata)) {
            XML_StopParser(data->xmlparser, XML_FALSE);
            return;
          }
        }
      } else {
        //for non-calback method suspend here with cell data
        if (!data->celldata)
          data->celldata = XML_Char_dup(X(""));
        XML_StopParser(data->xmlparser, XML_TRUE);
      }
    }
  } else {
    data_sheet_expat_callback_find_row_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_value_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("v")) == 0 || XML_Char_icmp_ins(name, X("t")) == 0) {
    XML_SetElementHandler(data->xmlparser, NULL, data_sheet_expat_callback_find_value_end);
    XML_SetCharacterDataHandler(data->xmlparser, data_sheet_expat_callback_value_data);
  } else if (XML_Char_icmp_ins(name, X("is")) == 0) {
    data->cell_string_type = inline_string;
  } else if (XML_Char_icmp_ins(name, X("rPh")) == 0) {
    data->skiptag = XML_Char_dup(name);
    data->skiptagcount = 1;
    data->skip_start = data_sheet_expat_callback_find_value_start;
    data->skip_end = data_sheet_expat_callback_find_cell_end;
    data->skip_data = NULL;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_skip_tag_start, data_sheet_expat_callback_skip_tag_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  }
}

void data_sheet_expat_callback_find_value_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (XML_Char_icmp_ins(name, X("v")) == 0 || XML_Char_icmp_ins(name, X("t")) == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_value_start, data_sheet_expat_callback_find_cell_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  } else if (XML_Char_icmp_ins(name, X("is")) == 0) {
    data->cell_string_type = none;
  } else {
    data_sheet_expat_callback_find_row_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_value_data (void* callbackdata, const XML_Char* buf, int buflen)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (data->cell_string_type != none) {
    if ((data->celldata = XML_Char_realloc(data->celldata, data->celldatalen + buflen + 1)) == NULL) {
      //memory allocation error
      data->celldatalen = 0;
    } else {
      //add new data to value buffer
      XML_Char_poscpy(data->celldata, data->celldatalen, buf, buflen);
      data->celldatalen += buflen;
    }
  }
}

////////////////////////////////////////////////////////////////////////

struct xlsxio_read_sheet_struct {
  xlsxioreader handle;
  ZIPFILEENTRYTYPE* zipfile;
  struct data_sheet_callback_data processcallbackdata;
  size_t lastrownr;
  size_t paddingrow;
  size_t lastcolnr;
  size_t paddingcol;
};

DLL_EXPORT_XLSXIO size_t xlsxioread_sheet_last_row_index (xlsxioreadersheet sheethandle)
{
  return sheethandle->lastrownr;
}

DLL_EXPORT_XLSXIO size_t xlsxioread_sheet_last_column_index (xlsxioreadersheet sheethandle)
{
  return sheethandle->lastcolnr;
}

DLL_EXPORT_XLSXIO unsigned int xlsxioread_sheet_flags (xlsxioreadersheet sheethandle)
{
  return sheethandle->processcallbackdata.flags;
}

DLL_EXPORT_XLSXIO int xlsxioread_process (xlsxioreader handle, const XLSXIOCHAR* sheetname, unsigned int flags, xlsxioread_process_cell_callback_fn cell_callback, xlsxioread_process_row_callback_fn row_callback, void* callbackdata)
{
  int result = 0;
  //determine sheet file name
  struct main_sheet_get_rels_callback_data getrelscallbackdata = {
    .sheetname = sheetname,
    .basepath = NULL,
    .sheetrelid = NULL,
    .sheetfile = NULL,
    .sharedstringsfile = NULL,
    .stylesfile = NULL
  };
  iterate_files_by_contenttype(handle->zip, xlsx_content_type, main_sheet_get_sheetfile_callback, &getrelscallbackdata, NULL);
  if (!getrelscallbackdata.sheetrelid)
    iterate_files_by_contenttype(handle->zip, xlsm_content_type, main_sheet_get_sheetfile_callback, &getrelscallbackdata, NULL);
  if (!getrelscallbackdata.sheetrelid)
    iterate_files_by_contenttype(handle->zip, xltx_content_type, main_sheet_get_sheetfile_callback, &getrelscallbackdata, NULL);
  if (!getrelscallbackdata.sheetrelid)
    iterate_files_by_contenttype(handle->zip, xltm_content_type, main_sheet_get_sheetfile_callback, &getrelscallbackdata, NULL);

  //process shared strings
  struct sharedstringlist* sharedstrings = NULL;
  if (getrelscallbackdata.sharedstringsfile && getrelscallbackdata.sharedstringsfile[0]) {
    sharedstrings = sharedstringlist_create();
    struct shared_strings_callback_data sharedstringsdata;
    shared_strings_callback_data_initialize(&sharedstringsdata, sharedstrings);
    if (expat_process_zip_file(handle->zip, getrelscallbackdata.sharedstringsfile, shared_strings_callback_find_sharedstringtable_start, NULL, NULL, &sharedstringsdata, &sharedstringsdata.xmlparser) != 0) {
      //no shared strings found
      sharedstringlist_destroy(sharedstrings);
      sharedstrings = NULL;
    }
    shared_strings_callback_data_cleanup(&sharedstringsdata);
  }

  //process sheet
  if (!(flags & XLSXIOREAD_NO_CALLBACK)) {
    //use callback mechanism
    struct data_sheet_callback_data processcallbackdata;
    data_sheet_callback_data_initialize(&processcallbackdata, sharedstrings, flags, cell_callback, row_callback, callbackdata);
    expat_process_zip_file(handle->zip, getrelscallbackdata.sheetfile, data_sheet_expat_callback_find_worksheet_start, NULL, NULL, &processcallbackdata, &processcallbackdata.xmlparser);
    data_sheet_callback_data_cleanup(&processcallbackdata);
  } else {
    //use simplified interface by suspending the XML parser when data is found
    xlsxioreadersheet sheethandle = (xlsxioreadersheet)callbackdata;
    data_sheet_callback_data_initialize(&sheethandle->processcallbackdata, sharedstrings, flags, NULL, NULL, sheethandle);
    if ((sheethandle->zipfile = XML_Char_openzip(sheethandle->handle->zip, getrelscallbackdata.sheetfile, 0)) == NULL) {
      result = 1;
    }
    if ((sheethandle->processcallbackdata.xmlparser = expat_process_zip_file_suspendable(sheethandle->zipfile, data_sheet_expat_callback_find_worksheet_start, NULL, NULL, &sheethandle->processcallbackdata)) == NULL) {
      result = 2;
    }
  }

  //clean up
  free(getrelscallbackdata.basepath);
  free(getrelscallbackdata.sheetrelid);
  free(getrelscallbackdata.sheetfile);
  free(getrelscallbackdata.sharedstringsfile);
  free(getrelscallbackdata.stylesfile);
  return result;
}

////////////////////////////////////////////////////////////////////////

struct xlsxio_read_sheetlist_struct {
  xlsxioreader handle;
  ZIPFILEENTRYTYPE* zipfile;
  struct main_sheet_list_callback_data sheetcallbackdata;
  XML_Parser xmlparser;
  XML_Char* nextsheetname;
};

int xlsxioread_list_sheets_resumable_callback (const XLSXIOCHAR* name, void* callbackdata)
{
  //struct main_sheet_list_callback_data* data = (struct main_sheet_list_callback_data*)callbackdata;
  xlsxioreadersheetlist data = (xlsxioreadersheetlist)callbackdata;
  data->nextsheetname = XML_Char_dup(name);
  XML_StopParser(data->xmlparser, XML_TRUE);
  return 0;
}

void xlsxioread_find_main_sheet_file_callback (ZIPFILETYPE* zip, const XML_Char* filename, const XML_Char* contenttype, void* callbackdata)
{
  XML_Char** data = (XML_Char**)callbackdata;
  *data = XML_Char_dup(filename);
}

DLL_EXPORT_XLSXIO xlsxioreadersheetlist xlsxioread_sheetlist_open (xlsxioreader handle)
{
  //determine main sheet name
  XML_Char* mainsheetfile = NULL;
  iterate_files_by_contenttype(handle->zip, xlsx_content_type, xlsxioread_find_main_sheet_file_callback, &mainsheetfile, NULL);
  if (!mainsheetfile)
    iterate_files_by_contenttype(handle->zip, xlsm_content_type, xlsxioread_find_main_sheet_file_callback, &mainsheetfile, NULL);
  if (!mainsheetfile)
    iterate_files_by_contenttype(handle->zip, xltx_content_type, xlsxioread_find_main_sheet_file_callback, &mainsheetfile, NULL);
  if (!mainsheetfile)
    iterate_files_by_contenttype(handle->zip, xltm_content_type, xlsxioread_find_main_sheet_file_callback, &mainsheetfile, NULL);
  if (!mainsheetfile)
    return NULL;
  //process contents of main sheet
  xlsxioreadersheetlist result;
  if ((result = (xlsxioreadersheetlist)malloc(sizeof(struct xlsxio_read_sheetlist_struct))) == NULL)
    return NULL;
  result->handle = handle;
  result->sheetcallbackdata.xmlparser = NULL;
  result->sheetcallbackdata.callback = xlsxioread_list_sheets_resumable_callback;
  result->sheetcallbackdata.callbackdata = result;
  result->nextsheetname = NULL;
  if ((result->zipfile = XML_Char_openzip(handle->zip, mainsheetfile, 0)) != NULL) {
    result->xmlparser = expat_process_zip_file_suspendable(result->zipfile, main_sheet_list_expat_callback_element_start, NULL, NULL, &result->sheetcallbackdata);
  }
  //clean up
  free(mainsheetfile);
  return result;
}

DLL_EXPORT_XLSXIO void xlsxioread_sheetlist_close (xlsxioreadersheetlist sheetlisthandle)
{
  if (!sheetlisthandle)
    return;
  if (sheetlisthandle->xmlparser)
    XML_ParserFree(sheetlisthandle->xmlparser);
  if (sheetlisthandle->zipfile)
#ifdef USE_MINIZIP
    unzCloseCurrentFile(sheetlisthandle->zipfile);
#else
    zip_fclose(sheetlisthandle->zipfile);
#endif
  free(sheetlisthandle->nextsheetname);
  free(sheetlisthandle);

}

DLL_EXPORT_XLSXIO const XLSXIOCHAR* xlsxioread_sheetlist_next (xlsxioreadersheetlist sheetlisthandle)
{
  if (!sheetlisthandle->zipfile || !sheetlisthandle->xmlparser)
    return NULL;
  free(sheetlisthandle->nextsheetname);
  sheetlisthandle->nextsheetname = NULL;
  enum XML_Status status;
  if ((status = expat_process_zip_file_resume(sheetlisthandle->zipfile, sheetlisthandle->xmlparser)) != XML_STATUS_SUSPENDED) {
    return NULL;
  }
  return sheetlisthandle->nextsheetname;
}

////////////////////////////////////////////////////////////////////////

DLL_EXPORT_XLSXIO xlsxioreadersheet xlsxioread_sheet_open (xlsxioreader handle, const XLSXIOCHAR* sheetname, unsigned int flags)
{
  xlsxioreadersheet result;
  if ((result = (xlsxioreadersheet)malloc(sizeof(struct xlsxio_read_sheet_struct))) == NULL)
    return NULL;
  result->handle = handle;
  result->zipfile = NULL;
  result->lastrownr = 0;
  result->paddingrow = 0;
  result->lastcolnr = 0;
  result->paddingcol = 0;
  xlsxioread_process(handle, sheetname, flags | XLSXIOREAD_NO_CALLBACK, NULL, NULL, result);
  return result;
}

DLL_EXPORT_XLSXIO void xlsxioread_sheet_close (xlsxioreadersheet sheethandle)
{
  if (!sheethandle)
    return;
  if (sheethandle->processcallbackdata.xmlparser)
    XML_ParserFree(sheethandle->processcallbackdata.xmlparser);
  data_sheet_callback_data_cleanup(&sheethandle->processcallbackdata);
  if (sheethandle->zipfile)
#ifdef USE_MINIZIP
    unzCloseCurrentFile(sheethandle->zipfile);
#else
    zip_fclose(sheethandle->zipfile);
#endif
  free(sheethandle);
}

DLL_EXPORT_XLSXIO int xlsxioread_sheet_next_row (xlsxioreadersheet sheethandle)
{
  enum XML_Status status;
  if (!sheethandle) {
    return 0;
  }
  sheethandle->lastcolnr = 0;
  //when padding rows don't retrieve new data
  if (sheethandle->paddingrow) {
    if (sheethandle->paddingrow < sheethandle->processcallbackdata.rownr) {
      return 3;
    } else {
      sheethandle->paddingrow = 0;
      return 2;
    }
  }
  sheethandle->paddingcol = 0;
  //go to beginning of next row
  while ((status = expat_process_zip_file_resume(sheethandle->zipfile, sheethandle->processcallbackdata.xmlparser)) == XML_STATUS_SUSPENDED && sheethandle->processcallbackdata.colnr != 0) {
  }
  return (status == XML_STATUS_SUSPENDED ? 1 : 0);
}

DLL_EXPORT_XLSXIO XLSXIOCHAR* xlsxioread_sheet_next_cell (xlsxioreadersheet sheethandle)
{
  XML_Char* result;
  if (!sheethandle)
    return NULL;
  //append empty column if needed
  if (sheethandle->paddingcol) {
    if (sheethandle->paddingcol > sheethandle->processcallbackdata.cols) {
      //last empty column added, finish row
      sheethandle->paddingcol = 0;
      //when padding rows prepare for the next one
      if (sheethandle->paddingrow) {
        sheethandle->lastrownr++;
        sheethandle->paddingrow++;
        if (sheethandle->paddingrow + 1 < sheethandle->processcallbackdata.rownr) {
          sheethandle->paddingcol = 1;
        }
      }
      return NULL;
    } else {
      //add another empty column
      sheethandle->paddingcol++;
      return XML_Char_dup(X(""));
    }
  }
  //get value
  if (!sheethandle->processcallbackdata.celldata)
    if (expat_process_zip_file_resume(sheethandle->zipfile, sheethandle->processcallbackdata.xmlparser) != XML_STATUS_SUSPENDED)
      sheethandle->processcallbackdata.celldata = NULL;
  //insert empty rows if needed
  if (!(sheethandle->processcallbackdata.flags & XLSXIOREAD_SKIP_EMPTY_ROWS) && sheethandle->lastrownr + 1 < sheethandle->processcallbackdata.rownr) {
    sheethandle->paddingrow = sheethandle->lastrownr + 1;
    sheethandle->paddingcol = sheethandle->processcallbackdata.colnr*0 + 1;
    return xlsxioread_sheet_next_cell(sheethandle);
  }
  //insert empty column before if needed
  if (!(sheethandle->processcallbackdata.flags & XLSXIOREAD_SKIP_EMPTY_CELLS)) {
    if (sheethandle->lastcolnr + 1 < sheethandle->processcallbackdata.colnr) {
      sheethandle->lastcolnr++;
      return XML_Char_dup(X(""));
    }
  }
  result = sheethandle->processcallbackdata.celldata;
  sheethandle->processcallbackdata.celldata = NULL;
  //end of row
  if (!result) {
    sheethandle->lastrownr = sheethandle->processcallbackdata.rownr;
    //insert empty column at end if row if needed
    if (!result && !(sheethandle->processcallbackdata.flags & XLSXIOREAD_SKIP_EMPTY_CELLS) && sheethandle->processcallbackdata.colnr < sheethandle->processcallbackdata.cols) {
      sheethandle->paddingcol = sheethandle->lastcolnr + 1;
      return xlsxioread_sheet_next_cell(sheethandle);
    }
  }
  sheethandle->lastcolnr = sheethandle->processcallbackdata.colnr;
  return result;
}

DLL_EXPORT_XLSXIO int xlsxioread_sheet_next_cell_string (xlsxioreadersheet sheethandle, XLSXIOCHAR** pvalue)
{
  XML_Char* result;
  if (!sheethandle)
    return -1;
  if ((result = xlsxioread_sheet_next_cell(sheethandle)) == NULL)
    return 0;
  if (pvalue)
    *pvalue = result;
  return 1;
}

DLL_EXPORT_XLSXIO int xlsxioread_sheet_next_cell_int (xlsxioreadersheet sheethandle, int64_t* pvalue)
{
  XML_Char* result;
  int status;
  if ((result = xlsxioread_sheet_next_cell(sheethandle)) == NULL)
    return 0;
  if (pvalue) {
    status = XML_Char_sscanf(result, X("%" PRIi64), pvalue);
    if (status == EOF || status == 0)
      *pvalue = 0;
    //alternative: use strtoimax()
  }
  free(result);
  return 1;
}

DLL_EXPORT_XLSXIO int xlsxioread_sheet_next_cell_float (xlsxioreadersheet sheethandle, double* pvalue)
{
  XML_Char* result;
  if ((result = xlsxioread_sheet_next_cell(sheethandle)) == NULL)
    return 0;
  if (pvalue)
    *pvalue = XML_Char_tod(result);
  free(result);
  return 1;
}

DLL_EXPORT_XLSXIO int xlsxioread_sheet_next_cell_datetime (xlsxioreadersheet sheethandle, time_t* pvalue)
{
  XML_Char* result;
  if ((result = xlsxioread_sheet_next_cell(sheethandle)) == NULL)
    return 0;
  if (pvalue) {
    double value = XML_Char_tod(result);
    if (value != 0) {
      value = (value - 25569) * 86400;  //converstion from Excel to Unix timestamp
    }
    *pvalue = value;
  }
  free(result);
  return 1;
}

