#include <stdlib.h>
#include <string.h>
#include <unistd.h>
#include <errno.h>

#include "ioapi.h"
#include "unzip.h"

#include "lxr_zip.h"

struct lxr_zip {
    unzFile uf;
    void   *mem_buf;
    size_t  mem_len;
    int     owns_mem;
};

struct lxr_zip_file {
    unzFile uf;
};

/* ------------------------------------------------------------------------- */
/* Memory IO callbacks for minizip                                           */
/* ------------------------------------------------------------------------- */

typedef struct {
    const unsigned char *base;
    size_t               size;
    size_t               pos;
} lxr_mem_stream;

static voidpf ZCALLBACK lxr_mem_open(voidpf opaque, const char *filename, int mode)
{
    (void)filename;
    (void)mode;
    return opaque;
}

static uLong ZCALLBACK lxr_mem_read(voidpf opaque, voidpf stream, void *buf, uLong size)
{
    lxr_mem_stream *s = (lxr_mem_stream *)stream;
    size_t avail;
    if (!s) return 0;
    avail = s->size - s->pos;
    if (avail > size) avail = size;
    memcpy(buf, s->base + s->pos, avail);
    s->pos += avail;
    return (uLong)avail;
}

static long ZCALLBACK lxr_mem_tell(voidpf opaque, voidpf stream)
{
    lxr_mem_stream *s = (lxr_mem_stream *)stream;
    return (long)s->pos;
}

static long ZCALLBACK lxr_mem_seek(voidpf opaque, voidpf stream, uLong offset, int origin)
{
    lxr_mem_stream *s = (lxr_mem_stream *)stream;
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

static int ZCALLBACK lxr_mem_close(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

static int ZCALLBACK lxr_mem_error(voidpf opaque, voidpf stream)
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
} lxr_fd_stream;

static voidpf ZCALLBACK lxr_fd_open(voidpf opaque, const char *filename, int mode)
{
    (void)filename;
    (void)mode;
    return opaque;
}

static uLong ZCALLBACK lxr_fd_read(voidpf opaque, voidpf stream, void *buf, uLong size)
{
    lxr_fd_stream *s = (lxr_fd_stream *)stream;
    ssize_t n = read(s->fd, buf, size);
    return n < 0 ? 0 : (uLong)n;
}

static long ZCALLBACK lxr_fd_tell(voidpf opaque, voidpf stream)
{
    lxr_fd_stream *s = (lxr_fd_stream *)stream;
    return (long)lseek(s->fd, 0, SEEK_CUR);
}

static long ZCALLBACK lxr_fd_seek(voidpf opaque, voidpf stream, uLong offset, int origin)
{
    lxr_fd_stream *s = (lxr_fd_stream *)stream;
    int whence;
    switch (origin) {
    case ZLIB_FILEFUNC_SEEK_SET: whence = SEEK_SET; break;
    case ZLIB_FILEFUNC_SEEK_CUR: whence = SEEK_CUR; break;
    case ZLIB_FILEFUNC_SEEK_END: whence = SEEK_END; break;
    default: return -1;
    }
    return lseek(s->fd, (off_t)offset, whence) < 0 ? -1 : 0;
}

static int ZCALLBACK lxr_fd_close(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

static int ZCALLBACK lxr_fd_error(voidpf opaque, voidpf stream)
{
    (void)opaque;
    (void)stream;
    return 0;
}

/* ------------------------------------------------------------------------- */
/* Public API                                                                */
/* ------------------------------------------------------------------------- */

lxr_zip *lxr_zip_open_path(const char *path)
{
    lxr_zip *z;
    if (!path) return NULL;
    z = (lxr_zip *)calloc(1, sizeof(*z));
    if (!z) return NULL;
    z->uf = unzOpen64(path);
    if (!z->uf) {
        free(z);
        return NULL;
    }
    return z;
}

lxr_zip *lxr_zip_open_memory(const void *data, size_t len)
{
    zlib_filefunc_def funcs;
    lxr_mem_stream   *stream;
    lxr_zip          *z;
    void             *buf;

    if (!data || len == 0) return NULL;

    /* Take a private copy so the caller's buffer can be freed independently. */
    buf = malloc(len);
    if (!buf) return NULL;
    memcpy(buf, data, len);

    stream = (lxr_mem_stream *)calloc(1, sizeof(*stream));
    if (!stream) {
        free(buf);
        return NULL;
    }
    stream->base = (const unsigned char *)buf;
    stream->size = len;
    stream->pos  = 0;

    funcs.zopen_file  = lxr_mem_open;
    funcs.zread_file  = lxr_mem_read;
    funcs.zwrite_file = NULL;
    funcs.ztell_file  = lxr_mem_tell;
    funcs.zseek_file  = lxr_mem_seek;
    funcs.zclose_file = lxr_mem_close;
    funcs.zerror_file = lxr_mem_error;
    funcs.opaque      = stream;

    z = (lxr_zip *)calloc(1, sizeof(*z));
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
    /* stream is referenced by minizip via opaque; freed in lxr_zip_close. */
    z->mem_buf = stream;  /* stash for free; original buf already owned by stream */
    /* preserve buf via stream->base; freed by us below */
    return z;
}

lxr_zip *lxr_zip_open_fd(int fd)
{
    zlib_filefunc_def funcs;
    lxr_fd_stream    *stream;
    lxr_zip          *z;

    if (fd < 0) return NULL;

    stream = (lxr_fd_stream *)calloc(1, sizeof(*stream));
    if (!stream) return NULL;
    stream->fd = fd;

    funcs.zopen_file  = lxr_fd_open;
    funcs.zread_file  = lxr_fd_read;
    funcs.zwrite_file = NULL;
    funcs.ztell_file  = lxr_fd_tell;
    funcs.zseek_file  = lxr_fd_seek;
    funcs.zclose_file = lxr_fd_close;
    funcs.zerror_file = lxr_fd_error;
    funcs.opaque      = stream;

    z = (lxr_zip *)calloc(1, sizeof(*z));
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

void lxr_zip_close(lxr_zip *z)
{
    if (!z) return;
    if (z->uf) unzClose(z->uf);
    if (z->mem_buf) {
        if (z->owns_mem) {
            /* mem_buf actually points to lxr_mem_stream; free the underlying base too */
            lxr_mem_stream *s = (lxr_mem_stream *)z->mem_buf;
            free((void *)s->base);
            free(s);
        } else {
            free(z->mem_buf);
        }
    }
    free(z);
}

int lxr_zip_entry_exists(lxr_zip *z, const char *name)
{
    if (!z || !name) return 0;
    return unzLocateFile(z->uf, name, 0) == UNZ_OK ? 1 : 0;
}

lxr_zip_file *lxr_zip_open_entry(lxr_zip *z, const char *name)
{
    lxr_zip_file *zf;
    if (!z || !name) return NULL;
    if (unzLocateFile(z->uf, name, 0) != UNZ_OK) return NULL;
    if (unzOpenCurrentFile(z->uf) != UNZ_OK)     return NULL;

    zf = (lxr_zip_file *)calloc(1, sizeof(*zf));
    if (!zf) {
        unzCloseCurrentFile(z->uf);
        return NULL;
    }
    zf->uf = z->uf;
    return zf;
}

ssize_t lxr_zip_read(lxr_zip_file *zf, void *buf, size_t n)
{
    int got;
    if (!zf || !buf) return -1;
    if (n == 0) return 0;
    got = unzReadCurrentFile(zf->uf, buf, (unsigned int)n);
    return got < 0 ? -1 : (ssize_t)got;
}

void lxr_zip_close_entry(lxr_zip_file *zf)
{
    if (!zf) return;
    unzCloseCurrentFile(zf->uf);
    free(zf);
}

lxr_error lxr_zip_iterate_entries(lxr_zip *z, lxr_zip_iter_fn fn, void *ud)
{
    int rc;
    if (!z || !fn) return LXR_ERROR_NULL_PARAMETER;

    rc = unzGoToFirstFile(z->uf);
    while (rc == UNZ_OK) {
        char name[1024];
        unz_file_info64 info;
        rc = unzGetCurrentFileInfo64(z->uf, &info, name, sizeof(name) - 1, NULL, 0, NULL, 0);
        if (rc != UNZ_OK) break;
        name[sizeof(name) - 1] = 0;
        if (fn(name, ud) != 0) return LXR_NO_ERROR;  /* iteration stop requested */
        rc = unzGoToNextFile(z->uf);
    }
    return LXR_NO_ERROR;
}
