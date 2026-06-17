Q=@
ifdef V
Q=
endif

CC ?= cc
AR ?= ar
ARFLAGS ?= rc
PYTHON ?= python3

SRC_DIR := ../src
READER_SRC_DIR := ../reader/src
INC_DIR := ../include
EXPAT_DIR := ../../libexpat/expat/lib
OBJ_DIR := .test_obj
LIB_TEST := $(SRC_DIR)/libxlsx_test.a

READER_TEST_DIR := reader
FIXTURES_DIR := fixtures

CFLAGS += -DTESTING -DCOLOR_OK -DXML_POOR_ENTROPY -DNOCRYPT -DNOUNCRYPT -g -Wall -Wextra -Wno-unused-parameter
CPPFLAGS += -I$(INC_DIR) -I$(READER_SRC_DIR) -I../third_party/minizip -I$(EXPAT_DIR) -I$(READER_TEST_DIR)
LIBS_O := -lz

WRITER_SRCS := $(wildcard $(SRC_DIR)/*.c)
READER_SRCS := $(wildcard $(READER_SRC_DIR)/*.c)
MINIZIP_SRCS := \
	../third_party/minizip/ioapi.c \
	../third_party/minizip/unzip.c \
	../third_party/minizip/zip.c
TMPFILEPLUS_SRCS := ../third_party/tmpfileplus/tmpfileplus.c
MD5_SRCS := ../third_party/md5/md5.c
EXPAT_SRCS := \
	$(EXPAT_DIR)/xmlparse.c \
	$(EXPAT_DIR)/xmlrole.c \
	$(EXPAT_DIR)/xmltok.c

WRITER_OBJS := $(patsubst $(SRC_DIR)/%.c,$(OBJ_DIR)/writer/%.o,$(WRITER_SRCS))
READER_OBJS := $(patsubst $(READER_SRC_DIR)/%.c,$(OBJ_DIR)/reader/%.o,$(READER_SRCS))
MINIZIP_OBJS := $(patsubst ../third_party/minizip/%.c,$(OBJ_DIR)/minizip/%.o,$(MINIZIP_SRCS))
TMPFILEPLUS_OBJS := $(patsubst ../third_party/tmpfileplus/%.c,$(OBJ_DIR)/tmpfileplus/%.o,$(TMPFILEPLUS_SRCS))
MD5_OBJS := $(patsubst ../third_party/md5/%.c,$(OBJ_DIR)/md5/%.o,$(MD5_SRCS))
EXPAT_OBJS := $(patsubst $(EXPAT_DIR)/%.c,$(OBJ_DIR)/expat/%.o,$(EXPAT_SRCS))
LIB_OBJS := $(WRITER_OBJS) $(READER_OBJS) $(MINIZIP_OBJS) $(TMPFILEPLUS_OBJS) $(MD5_OBJS) $(EXPAT_OBJS)

READER_TEST_SRCS := $(wildcard $(READER_TEST_DIR)/test_*.c)
READER_TEST_BINS := $(patsubst $(READER_TEST_DIR)/%.c,$(READER_TEST_DIR)/%,$(READER_TEST_SRCS))
READER_FIXTURES := \
	$(FIXTURES_DIR)/types.xlsx \
	$(FIXTURES_DIR)/phase1.xlsx \
	$(FIXTURES_DIR)/phase3.xlsx \
	$(FIXTURES_DIR)/phase4.xlsx

.PHONY: all test clean writer-test reader-test reader-fixtures

all: $(LIB_TEST)

test: writer-test reader-test

writer-test: $(LIB_TEST)
	$(Q)$(MAKE) -C writer test \
		LIBS_A="$(abspath $(LIB_TEST))" \
		INC_DIR="$(abspath $(INC_DIR))" \
		LIB_DIR="$(abspath $(SRC_DIR))"

reader-test: reader-fixtures $(READER_TEST_BINS)
	$(Q)set -e; for test_bin in $(READER_TEST_BINS); do ./$$test_bin; done

reader-fixtures: $(READER_FIXTURES)

$(LIB_TEST): $(LIB_OBJS)
	$(Q)$(AR) $(ARFLAGS) $@ $^

$(OBJ_DIR)/writer $(OBJ_DIR)/reader $(OBJ_DIR)/minizip $(OBJ_DIR)/tmpfileplus $(OBJ_DIR)/md5 $(OBJ_DIR)/expat:
	$(Q)mkdir -p $@

$(OBJ_DIR)/writer/%.o: $(SRC_DIR)/%.c | $(OBJ_DIR)/writer
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/reader/%.o: $(READER_SRC_DIR)/%.c | $(OBJ_DIR)/reader
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/minizip/%.o: ../third_party/minizip/%.c | $(OBJ_DIR)/minizip
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/tmpfileplus/%.o: ../third_party/tmpfileplus/%.c | $(OBJ_DIR)/tmpfileplus
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/md5/%.o: ../third_party/md5/%.c | $(OBJ_DIR)/md5
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/expat/%.o: $(EXPAT_DIR)/%.c | $(OBJ_DIR)/expat
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(READER_TEST_DIR)/%: $(READER_TEST_DIR)/%.c $(LIB_TEST) $(READER_TEST_DIR)/unity.h $(READER_TEST_DIR)/lxlsx_reader_test_paths.h
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIB_TEST) $(LIBS_O)

$(FIXTURES_DIR)/types.xlsx: $(FIXTURES_DIR)/gen_types_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase1.xlsx: $(FIXTURES_DIR)/gen_phase1_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase3.xlsx: $(FIXTURES_DIR)/gen_phase3_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase4.xlsx: $(FIXTURES_DIR)/gen_phase4_xlsx.py
	$(Q)$(PYTHON) $< $@

clean:
	$(Q)$(MAKE) -C writer clean \
		LIBS_A="$(abspath $(LIB_TEST))" \
		INC_DIR="$(abspath $(INC_DIR))" \
		LIB_DIR="$(abspath $(SRC_DIR))"
	$(Q)rm -rf $(OBJ_DIR) $(LIB_TEST) $(READER_TEST_BINS) $(READER_FIXTURES) $(READER_TEST_DIR)/*.dSYM
