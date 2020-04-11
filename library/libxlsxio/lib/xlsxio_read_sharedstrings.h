#ifndef INCLUDED_XLSXIO_READ_SHAREDSTRINGS_H
#define INCLUDED_XLSXIO_READ_SHAREDSTRINGS_H

#include <stdint.h>
#include <expat.h>

#ifdef __cplusplus
extern "C" {
#endif

struct sharedstringlist {
  XML_Char** strings;
  size_t numstrings;
};

struct sharedstringlist* sharedstringlist_create ();
void sharedstringlist_destroy (struct sharedstringlist* sharedstrings);
size_t sharedstringlist_size (struct sharedstringlist* sharedstrings);
int sharedstringlist_add_buffer (struct sharedstringlist* sharedstrings, const XML_Char* data, size_t datalen);
int sharedstringlist_add_string (struct sharedstringlist* sharedstrings, const XML_Char* data);
const XML_Char* sharedstringlist_get (struct sharedstringlist* sharedstrings, size_t index);

////////////////////////////////////////////////////////////////////////

struct shared_strings_callback_data {
  XML_Parser xmlparser;
  struct sharedstringlist* sharedstrings;
  int insst;
  int insi;
  int intext;
  XML_Char* text;
  size_t textlen;
  XML_Char* skiptag;                    //tag to skip
  size_t skiptagcount;                  //nesting level for current tag to skip
  XML_StartElementHandler skip_start;   //start handler to set after skipping
  XML_EndElementHandler skip_end;       //end handler to set after skipping
  XML_CharacterDataHandler skip_data;   //data handler to set after skipping
};

void shared_strings_callback_data_initialize (struct shared_strings_callback_data* data, struct sharedstringlist* sharedstrings);
void shared_strings_callback_data_cleanup (struct shared_strings_callback_data* data);
void shared_strings_callback_skip_tag_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_skip_tag_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_find_sharedstringtable_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_sharedstringtable_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_find_shared_stringitem_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_shared_stringitem_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_find_shared_string_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_shared_string_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_string_data (void* callbackdata, const XML_Char* buf, int buflen);

#ifdef __cplusplus
}
#endif

#endif
