/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension                                                  |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2018 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <dev@service.viest.me>                                 |
  +----------------------------------------------------------------------+
*/

#include "xlswriter.h"
#include "php_streams.h"
#include "ext/standard/file.h"

/* {{{ */
unsigned int xlsx_to_csv(
        zval *stream_resource,
        const char *delimiter_str, int delimiter_str_len,
        const char *enclosure_str, int enclosure_str_len,
        const char *escape_str, int escape_str_len,
        xlsxioreadersheet sheet_t,
        zval *zv_type_arr_t, zend_long data_type_default,
        unsigned int flag, zend_fcall_info *fci, zend_fcall_info_cache *fci_cache
)
{
    ssize_t ret = 0;
    zval *_zv_type_arr_t = NULL;
    php_stream *_stream_t = NULL;
    char delimiter = ',', enclosure = '"', escape_char = '\\';

    ZEND_ASSERT(Z_TYPE_P(stream_resource) == IS_RESOURCE);

    if (((_stream_t) = (php_stream *)zend_fetch_resource2((Z_RES_P(stream_resource)),
            "stream", php_file_le_stream(), php_file_le_pstream())) == NULL) {
        return XLSWRITER_FALSE;
	}

    if (delimiter_str != NULL) {
        if (delimiter_str_len < 1) {
            zend_throw_exception(vtiful_exception_ce, "delimiter must be a character", 190);
            return XLSWRITER_FALSE;
        } else if (delimiter_str_len > 1) {
            zend_throw_exception(vtiful_exception_ce, "delimiter must be a single character", 191);
            return XLSWRITER_FALSE;
        }

        delimiter = *delimiter_str;
    }

    if (enclosure_str != NULL) {
        if (enclosure_str_len < 1) {
            zend_throw_exception(vtiful_exception_ce, "enclosure must be a character", 192);
            return XLSWRITER_FALSE;
        } else if (enclosure_str_len > 1) {
            zend_throw_exception(vtiful_exception_ce, "enclosure must be a single character", 193);
            return XLSWRITER_FALSE;
        }

        enclosure = *enclosure_str;
    }

    if (escape_str != NULL) {
        if (escape_str_len < 1) {
            zend_throw_exception(vtiful_exception_ce, "escape must be a character", 194);
            return XLSWRITER_FALSE;
        } else if (escape_str_len > 1) {
            zend_throw_exception(vtiful_exception_ce, "escape must be a single character", 195);
            return XLSWRITER_FALSE;
        }

        escape_char = *escape_str;
    }

    if (Z_TYPE_P(zv_type_arr_t) == IS_ARRAY) {
        _zv_type_arr_t = zv_type_arr_t;
    }

    zval _zv_tmp_row;
    ZVAL_NULL(&_zv_tmp_row);

    while (sheet_read_row(sheet_t))
    {
        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, _zv_type_arr_t, data_type_default, flag);

        if (fci != NULL && fci_cache != NULL) {
            zval retval;

            fci->retval      = &retval;
            fci->params      = &_zv_tmp_row;
            fci->param_count = 1;

            zend_call_function(fci, fci_cache);

            if (Z_TYPE(retval) == IS_ARRAY) {
#if PHP_VERSION_ID >= 80100
                ret = php_fputcsv(_stream_t, &retval, delimiter, enclosure, escape_char, NULL);
#else
                ret = php_fputcsv(_stream_t, &retval, delimiter, enclosure, escape_char);
#endif
            }

            zval_ptr_dtor(&retval);
            goto CLEAN_UP_SCENE;
        }

#if PHP_VERSION_ID >= 80100
        ret = php_fputcsv(_stream_t, &_zv_tmp_row, delimiter, enclosure, escape_char, NULL);
#else
        ret = php_fputcsv(_stream_t, &_zv_tmp_row, delimiter, enclosure, escape_char);
#endif

        CLEAN_UP_SCENE:

        zend_hash_clean(Z_ARRVAL(_zv_tmp_row));

        if (ret < 0) {
            return XLSWRITER_FALSE;
        }
    }

    zval_dtor(&_zv_tmp_row);

    return XLSWRITER_TRUE;
}
/* }}} */