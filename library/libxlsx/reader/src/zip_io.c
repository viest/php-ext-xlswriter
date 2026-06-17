#include <stdlib.h>
#include <string.h>
#include <errno.h>

#include "ioapi.h"
#include "unzip.h"

#include "lxlsx_reader_compat.h"
#include "lxlsx_reader_zip.h"

struct lxlsx_reader_zip {
    unzFile uf;
    void   *mem_buf;
    size_t  mem_len;
    int     owns_mem;
};

struct lxlsx_reader_zip_file {
    unzFile uf;
};

/* ------------------------------------------------------------------------- */
/* Memory IO callbacks for minizip                                           */
/* ------------------------------------------------------------------------- */

typedef struct {
    const unsigned char *base;
    size_t               size;
    size_t               pos;
} lxlsx_reader_mem_stream;

static voidpf ZCALLBACK lxlsx_reader_mem_open(voidpf opaque, const char *filename, int mode)
{
    (void)filename;
    (void)mode;
    return opaque;
}

static uLong ZCALLBACK lxlsx_reader_mem_read(voidpf opaque, voidpf stream, void *buf, uLong size)
{
    lxlsx_reader_mem_stream *s = (lxlsx_reader_mem_stream *)stream;
    size_t avail;
    if (!s) return 0;
    avail = s->size - s->pos;
    if (avail > size) avail = size;
    memcpy(buf, s->base + s->pos, avail);
    s->pos += avail;
    return (uLong)avail;
}

static long ZCALLBACK lxlsx_reader_mem_tell(voidpf opaque, voidpf stream)
{
    lxlsx_reader_mem_stream *s = (lxlsx_reader_mem_stream *)stream;
    return (long)s->pos;
}

static long ZCALLBACK lxlsx_reader_mem_seek(voidpf opaque, voidpf stream, uLong offset, int origin)
{
    lxlsx_reader_mem_stream *s = (lxlsx_reader_mem_stream *)stream;
    size_t want;
    switch (origin) {
    case ZLIB_FILEFUNC_SEEK_SET: want = offset;            break;
    case ZLIB_FILEFUNC_SEEK_CUR: want = s->pos + offset;   break;
    case ZLIB_FILEFUNC_SEEK_END: want = s->size + offset;  break;
    default: return -1;
    }
    if (want > s->size) return -1;
    s->pos = want;
    return 0;
}

static int ZCALLBACK lxlsx_reader_mem_close(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

static int ZCALLBACK lxlsx_reader_mem_error(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

/* ------------------------------------------------------------------------- */
/* fd IO callbacks for minizip                                               */
/* ------------------------------------------------------------------------- */

typedef struct {
    int fd;
} lxlsx_reader_fd_stream;

static voidpf ZCALLBACK lxlsx_reader_fd_open(voidpf opaque, const char *filename, int mode)
{
    (void)filename;
    (void)mode;
    return opaque;
}

static uLong ZCALLBACK lxlsx_reader_fd_read(voidpf opaque, voidpf stream, void *buf, uLong size)
{
    lxlsx_reader_fd_stream *s = (lxlsx_reader_fd_stream *)stream;
    ssize_t n = lxlsx_reader_read(s->fd, buf, size);
    return n < 0 ? 0 : (uLong)n;
}

static long ZCALLBACK lxlsx_reader_fd_tell(voidpf opaque, voidpf stream)
{
    lxlsx_reader_fd_stream *s = (lxlsx_reader_fd_stream *)stream;
    return (long)lxlsx_reader_lseek(s->fd, 0, SEEK_CUR);
}

static long ZCALLBACK lxlsx_reader_fd_seek(voidpf opaque, voidpf stream, uLong offset, int origin)
{
    lxlsx_reader_fd_stream *s = (lxlsx_reader_fd_stream *)stream;
    int whence;
    switch (origin) {
    case ZLIB_FILEFUNC_SEEK_SET: whence = SEEK_SET; break;
    case ZLIB_FILEFUNC_SEEK_CUR: whence = SEEK_CUR; break;
    case ZLIB_FILEFUNC_SEEK_END: whence = SEEK_END; break;
    default: return -1;
    }
    return lxlsx_reader_lseek(s->fd, (lxlsx_reader_off_t)offset, whence) < 0 ? -1 : 0;
}

static int ZCALLBACK lxlsx_reader_fd_close(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

static int ZCALLBACK lxlsx_reader_fd_error(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

/* ------------------------------------------------------------------------- */
/* Public API                                                                */
/* ------------------------------------------------------------------------- */

lxlsx_reader_zip *lxlsx_reader_zip_open_path(const char *path)
{
    lxlsx_reader_zip *z;
    if (!path) return NULL;
    z = (lxlsx_reader_zip *)calloc(1, sizeof(*z));
    if (!z) return NULL;
    z->uf = unzOpen64(path);
    if (!z->uf) {
        free(z);
        return NULL;
    }
    return z;
}

lxlsx_reader_zip *lxlsx_reader_zip_open_memory(const void *data, size_t len)
{
    zlib_filefunc_def funcs;
    lxlsx_reader_mem_stream   *stream;
    lxlsx_reader_zip          *z;
    void             *buf;

    if (!data || len == 0) return NULL;

    /* Take a private copy so the caller's buffer can be freed independently. */
    buf = malloc(len);
    if (!buf) return NULL;
    memcpy(buf, data, len);

    stream = (lxlsx_reader_mem_stream *)calloc(1, sizeof(*stream));
    if (!stream) {
        free(buf);
        return NULL;
    }
    stream->base = (const unsigned char *)buf;
    stream->size = len;
    stream->pos  = 0;

    funcs.zopen_file  = lxlsx_reader_mem_open;
    funcs.zread_file  = lxlsx_reader_mem_read;
    funcs.zwrite_file = NULL;
    funcs.ztell_file  = lxlsx_reader_mem_tell;
    funcs.zseek_file  = lxlsx_reader_mem_seek;
    funcs.zclose_file = lxlsx_reader_mem_close;
    funcs.zerror_file = lxlsx_reader_mem_error;
    funcs.opaque      = stream;

    z = (lxlsx_reader_zip *)calloc(1, sizeof(*z));
    if (!z) {
        free(stream);
        free(buf);
        return NULL;
    }
    z->mem_buf  = buf;
    z->mem_len  = len;
    z->owns_mem = 1;

    z->uf = unzOpen2(NULL, &funcs);
    if (!z->uf) {
        free(z);
        free(stream);
        free(buf);
        return NULL;
    }
    /* stream is referenced by minizip via opaque; freed in lxlsx_reader_zip_close. */
    z->mem_buf = stream;  /* stash for free; original buf already owned by stream */
    /* preserve buf via stream->base; freed by us below */
    return z;
}

lxlsx_reader_zip *lxlsx_reader_zip_open_fd(int fd)
{
    zlib_filefunc_def funcs;
    lxlsx_reader_fd_stream    *stream;
    lxlsx_reader_zip          *z;

    if (fd < 0) return NULL;

    stream = (lxlsx_reader_fd_stream *)calloc(1, sizeof(*stream));
    if (!stream) return NULL;
    stream->fd = fd;

    funcs.zopen_file  = lxlsx_reader_fd_open;
    funcs.zread_file  = lxlsx_reader_fd_read;
    funcs.zwrite_file = NULL;
    funcs.ztell_file  = lxlsx_reader_fd_tell;
    funcs.zseek_file  = lxlsx_reader_fd_seek;
    funcs.zclose_file = lxlsx_reader_fd_close;
    funcs.zerror_file = lxlsx_reader_fd_error;
    funcs.opaque      = stream;

    z = (lxlsx_reader_zip *)calloc(1, sizeof(*z));
    if (!z) {
        free(stream);
        return NULL;
    }
    z->mem_buf  = stream;
    z->mem_len  = 0;
    z->owns_mem = 0;

    z->uf = unzOpen2(NULL, &funcs);
    if (!z->uf) {
        free(z);
        free(stream);
        return NULL;
    }
    return z;
}

void lxlsx_reader_zip_close(lxlsx_reader_zip *z)
{
    if (!z) return;
    if (z->uf) unzClose(z->uf);
    if (z->mem_buf) {
        if (z->owns_mem) {
            /* mem_buf actually points to lxlsx_reader_mem_stream; free the underlying base too */
            lxlsx_reader_mem_stream *s = (lxlsx_reader_mem_stream *)z->mem_buf;
            free((void *)s->base);
            free(s);
        } else {
            free(z->mem_buf);
        }
    }
    free(z);
}

int lxlsx_reader_zip_entry_exists(lxlsx_reader_zip *z, const char *name)
{
    if (!z || !name) return 0;
    return unzLocateFile(z->uf, name, 0) == UNZ_OK ? 1 : 0;
}

lxlsx_reader_zip_file *lxlsx_reader_zip_open_entry(lxlsx_reader_zip *z, const char *name)
{
    lxlsx_reader_zip_file *zf;
    if (!z || !name) return NULL;
    if (unzLocateFile(z->uf, name, 0) != UNZ_OK) return NULL;
    if (unzOpenCurrentFile(z->uf) != UNZ_OK)     return NULL;

    zf = (lxlsx_reader_zip_file *)calloc(1, sizeof(*zf));
    if (!zf) {
        unzCloseCurrentFile(z->uf);
        return NULL;
    }
    zf->uf = z->uf;
    return zf;
}

ssize_t lxlsx_reader_zip_read(lxlsx_reader_zip_file *zf, void *buf, size_t n)
{
    int got;
    if (!zf || !buf) return -1;
    if (n == 0) return 0;
    got = unzReadCurrentFile(zf->uf, buf, (unsigned int)n);
    return got < 0 ? -1 : (ssize_t)got;
}

void lxlsx_reader_zip_close_entry(lxlsx_reader_zip_file *zf)
{
    if (!zf) return;
    unzCloseCurrentFile(zf->uf);
    free(zf);
}

lxlsx_reader_error lxlsx_reader_zip_iterate_entries(lxlsx_reader_zip *z, lxlsx_reader_zip_iter_fn fn, void *ud)
{
    int rc;
    if (!z || !fn) return LXLSX_READER_ERROR_NULL_PARAMETER;

    rc = unzGoToFirstFile(z->uf);
    while (rc == UNZ_OK) {
        char name[1024];
        unz_file_info64 info;
        rc = unzGetCurrentFileInfo64(z->uf, &info, name, sizeof(name) - 1, NULL, 0, NULL, 0);
        if (rc != UNZ_OK) break;
        name[sizeof(name) - 1] = 0;
        if (fn(name, ud) != 0) return LXLSX_READER_NO_ERROR;  /* iteration stop requested */
        rc = unzGoToNextFile(z->uf);
    }
    return LXLSX_READER_NO_ERROR;
}
