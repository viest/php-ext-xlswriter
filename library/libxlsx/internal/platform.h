#ifndef LXLSX_PLATFORM_H
#define LXLSX_PLATFORM_H

/*
 * Portability shims for the libxlsx reader module. Private - keeps Windows-specific
 * headers out of the public API.
 *
 * - ssize_t: POSIX-only on Linux/macOS; on MSVC use SSIZE_T from BaseTsd.h.
 * - read / lseek wrappers: POSIX on Unix, _read / _lseeki64 on MSVC.
 * - off_t: synthesised on MSVC where the type is not declared in this scope.
 */

#include <stddef.h>

#if defined(_MSC_VER)
#  include <io.h>
#  include <BaseTsd.h>
   typedef SSIZE_T ssize_t;
   typedef long long lxlsx_reader_off_t;
#  define lxlsx_reader_read(fd, buf, n)        _read((fd), (buf), (unsigned int)(n))
#  define lxlsx_reader_lseek(fd, off, whence)  _lseeki64((fd), (off), (whence))
#else
#  include <sys/types.h>
#  include <unistd.h>
   typedef off_t lxlsx_reader_off_t;
#  define lxlsx_reader_read(fd, buf, n)        read((fd), (buf), (n))
#  define lxlsx_reader_lseek(fd, off, whence)  lseek((fd), (off), (whence))
#endif

#endif
