#include "xlsxio_private.h"
#include "xlsxio_read_sharedstrings.h"
#include <stdlib.h>
//#include <inttypes.h>
#include <string.h>

#if defined(_MSC_VER) || (defined(__MINGW32__) && !defined(__MINGW64__))
#define strcasecmp _stricmp
#endif
#ifdef _WIN32
#define wcscasecmp _wcsicmp
#endif

struct sharedstringlist* sharedstringlist_create ()
{
  struct sharedstringlist* sharedstrings;
  if ((sharedstrings = (struct sharedstringlist*)malloc(sizeof(struct sharedstringlist))) != NULL) {
    sharedstrings->strings = NULL;
    sharedstrings->numstrings = 0;
  }
  return sharedstrings;
}

void sharedstringlist_destroy (struct sharedstringlist* sharedstrings)
{
  if (sharedstrings) {
    size_t i;
    for (i = 0; i < sharedstrings->numstrings; i++)
      free(sharedstrings->strings[i]);
    free(sharedstrings->strings);
    free(sharedstrings);
  }
}

size_t sharedstringlist_size (struct sharedstringlist* sharedstrings)
{
  if (!sharedstrings)
    return 0;
  return sharedstrings->numstrings;
}

int sharedstringlist_add_buffer (struct sharedstringlist* sharedstrings, const XML_Char* data, size_t datalen)
{
  XML_Char* s;
  XML_Char** p;
  if (!sharedstrings)
    return 1;
  if (!data) {
    s = NULL;
  } else {
    if ((s = XML_Char_malloc(datalen + 1)) == NULL)
      return 2;
    XML_Char_poscpy(s, 0, data, datalen);
    s[datalen] = 0;
  }
  if ((p = (XML_Char**)realloc(sharedstrings->strings, (sharedstrings->numstrings + 1) * sizeof(sharedstrings->strings[0]))) == NULL) {
    free(s);
    return 3;
  }
  sharedstrings->strings = p;
  sharedstrings->strings[sharedstrings->numstrings++] = s;
  return 0;
}

int sharedstringlist_add_string (struct sharedstringlist* sharedstrings, const XML_Char* data)
{
  return sharedstringlist_add_buffer(sharedstrings, data, (data ? XML_Char_len(data) : 0));
}

const XML_Char* sharedstringlist_get (struct sharedstringlist* sharedstrings, size_t index)
{
  if (!sharedstrings || index >= sharedstrings->numstrings)
    return NULL;
  return sharedstrings->strings[index];
}

////////////////////////////////////////////////////////////////////////

void shared_strings_callback_data_initialize (struct shared_strings_callback_data* data, struct sharedstringlist* sharedstrings)
{
  data->xmlparser = NULL;
  data->sharedstrings = sharedstrings;
  data->insst = 0;
  data->insi = 0;
  data->intext = 0;
  data->text = NULL;
  data->textlen = 0;
  data->skiptag = NULL;
  data->skiptagcount = 0;
  data->skip_start = NULL;
  data->skip_end = NULL;
  data->skip_data = NULL;
}

void shared_strings_callback_data_cleanup (struct shared_strings_callback_data* data)
{
  free(data->text);
  free(data->skiptag);
}

void shared_strings_callback_skip_tag_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (name && XML_Char_icmp(name, data->skiptag) == 0) {
    //increment nesting level
    data->skiptagcount++;
  }
}

void shared_strings_callback_skip_tag_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (!name || XML_Char_icmp(name, data->skiptag) == 0) {
    if (--data->skiptagcount == 0) {
      //restore handlers when done skipping
      XML_SetElementHandler(data->xmlparser, data->skip_start, data->skip_end);
      XML_SetCharacterDataHandler(data->xmlparser, data->skip_data);
      free(data->skiptag);
      data->skiptag = NULL;
    }
  }
}

void shared_strings_callback_find_sharedstringtable_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("sst")) == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_stringitem_start, NULL);
  }
}

void shared_strings_callback_find_sharedstringtable_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("sst")) == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_sharedstringtable_start, NULL);
  }
}

void shared_strings_callback_find_shared_stringitem_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("si")) == 0) {
    if (data->text)
      free(data->text);
    data->text = NULL;
    data->textlen = 0;
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_string_start, shared_strings_callback_find_sharedstringtable_end);
  }
}

void shared_strings_callback_find_shared_stringitem_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("si")) == 0) {
    sharedstringlist_add_buffer(data->sharedstrings, data->text, data->textlen);
    if (data->text)
      free(data->text);
    data->text = NULL;
    data->textlen = 0;
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_stringitem_start, shared_strings_callback_find_sharedstringtable_end);
  } else {
    shared_strings_callback_find_sharedstringtable_end(callbackdata, name);
  }
}

void shared_strings_callback_find_shared_string_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("t")) == 0) {
    XML_SetElementHandler(data->xmlparser, NULL, shared_strings_callback_find_shared_string_end);
    XML_SetCharacterDataHandler(data->xmlparser, shared_strings_callback_string_data);
  } else if (XML_Char_icmp(name, X("rPh")) == 0) {
    data->skiptag = XML_Char_dup(name);
    data->skiptagcount = 1;
    data->skip_start = shared_strings_callback_find_shared_string_start;
    data->skip_end = shared_strings_callback_find_shared_stringitem_end;
    data->skip_data = NULL;
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_skip_tag_start, shared_strings_callback_skip_tag_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  }
}

void shared_strings_callback_find_shared_string_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (XML_Char_icmp(name, X("t")) == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_string_start, shared_strings_callback_find_shared_stringitem_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  } else {
    shared_strings_callback_find_shared_stringitem_end(callbackdata, name);
  }
}

void shared_strings_callback_string_data (void* callbackdata, const XML_Char* buf, int buflen)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if ((data->text = XML_Char_realloc(data->text, data->textlen + buflen)) == NULL) {
    //memory allocation error
    data->textlen = 0;
  } else {
    XML_Char_poscpy(data->text, data->textlen, buf, buflen);
    data->textlen += buflen;
  }
}

////////////////////////////////////////////////////////////////////////

