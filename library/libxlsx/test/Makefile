Q=@
ifdef V
Q=
endif

CC ?= cc
AR ?= ar
ARFLAGS ?= rc
PYTHON ?= python3

SRC_DIR := ../src
INC_DIR := ../include
EXPAT_DIR := ../../libexpat/expat/lib
OBJ_DIR := .test_obj
LIB_TEST := $(SRC_DIR)/libxlsx_test.a

READER_TEST_DIR := reader
INTEGRATION_TEST_DIR := integration
EDIT_TEST_DIR := edit
FIXTURES_DIR := fixtures

CFLAGS += -DTESTING -DCOLOR_OK -DXML_POOR_ENTROPY -DNOCRYPT -DNOUNCRYPT -g -Wall -Wextra -Wno-unused-parameter
CPPFLAGS += -I$(INC_DIR) -I$(SRC_DIR) -I../internal -I../third_party/minizip -I$(EXPAT_DIR) -I$(READER_TEST_DIR)
LIBS_O := -lz

LIBXLSX_SRCS := $(wildcard $(SRC_DIR)/*.c)
MINIZIP_SRCS := \
	../third_party/minizip/ioapi.c \
	../third_party/minizip/unzip.c \
	../third_party/minizip/zip.c
TMPFILEPLUS_SRCS := ../third_party/tmpfileplus/tmpfileplus.c
MD5_SRCS := ../third_party/md5/md5.c
DTOA_SRCS := ../third_party/dtoa/emyg_dtoa.c
EXPAT_SRCS := \
	$(EXPAT_DIR)/loadlibrary.c \
	$(EXPAT_DIR)/xmlparse.c \
	$(EXPAT_DIR)/xmlrole.c \
	$(EXPAT_DIR)/xmltok.c

LIBXLSX_OBJS := $(patsubst $(SRC_DIR)/%.c,$(OBJ_DIR)/libxlsx/%.o,$(LIBXLSX_SRCS))
MINIZIP_OBJS := $(patsubst ../third_party/minizip/%.c,$(OBJ_DIR)/minizip/%.o,$(MINIZIP_SRCS))
TMPFILEPLUS_OBJS := $(patsubst ../third_party/tmpfileplus/%.c,$(OBJ_DIR)/tmpfileplus/%.o,$(TMPFILEPLUS_SRCS))
MD5_OBJS := $(patsubst ../third_party/md5/%.c,$(OBJ_DIR)/md5/%.o,$(MD5_SRCS))
DTOA_OBJS :=
EXPAT_OBJS := $(patsubst $(EXPAT_DIR)/%.c,$(OBJ_DIR)/expat/%.o,$(EXPAT_SRCS))
ifdef USE_DTOA_LIBRARY
CFLAGS += -DUSE_DTOA_LIBRARY
DTOA_OBJS := $(patsubst ../third_party/dtoa/%.c,$(OBJ_DIR)/dtoa/%.o,$(DTOA_SRCS))
endif
LIB_OBJS := $(LIBXLSX_OBJS) $(MINIZIP_OBJS) $(TMPFILEPLUS_OBJS) $(MD5_OBJS) $(EXPAT_OBJS)
LIB_OBJS += $(DTOA_OBJS)

WRITER_TEST_DIR := writer
WRITER_TEST_BIN := $(WRITER_TEST_DIR)/test_all
WRITER_TEST_SRCS := \
	$(WRITER_TEST_DIR)/test_all.c \
	$(wildcard $(WRITER_TEST_DIR)/*/test*.c)
WRITER_TEST_OBJS := $(patsubst $(WRITER_TEST_DIR)/%.c,$(OBJ_DIR)/writer/%.o,$(WRITER_TEST_SRCS))
READER_TEST_SRCS := $(wildcard $(READER_TEST_DIR)/test_*.c)
READER_TEST_BINS := $(patsubst $(READER_TEST_DIR)/%.c,$(READER_TEST_DIR)/%,$(READER_TEST_SRCS))
INTEGRATION_TEST_SRCS := $(wildcard $(INTEGRATION_TEST_DIR)/test_*.c)
INTEGRATION_TEST_BINS := $(patsubst $(INTEGRATION_TEST_DIR)/%.c,$(INTEGRATION_TEST_DIR)/%,$(INTEGRATION_TEST_SRCS))
EDIT_TEST_SRCS := $(wildcard $(EDIT_TEST_DIR)/test_*.c)
EDIT_TEST_BINS := $(patsubst $(EDIT_TEST_DIR)/%.c,$(EDIT_TEST_DIR)/%,$(EDIT_TEST_SRCS))
FORMULA_TEST_DIR := formula
FORMULA_TEST_SRCS := $(wildcard $(FORMULA_TEST_DIR)/test_*.c)
FORMULA_TEST_BINS := $(patsubst $(FORMULA_TEST_DIR)/%.c,$(FORMULA_TEST_DIR)/%,$(FORMULA_TEST_SRCS))
INTEGRATION_OUTPUTS := $(FIXTURES_DIR)/integration_roundtrip.xlsx
EDIT_OUTPUTS := \
	$(FIXTURES_DIR)/edit_source.xlsx \
	$(FIXTURES_DIR)/edit_noop.xlsx \
	$(FIXTURES_DIR)/edit_number.xlsx \
	$(FIXTURES_DIR)/edit_namespace.xlsx \
	$(FIXTURES_DIR)/edit_self_closing.xlsx \
	$(FIXTURES_DIR)/edit_ordered_rows.xlsx \
	$(FIXTURES_DIR)/edit_snapshot.xlsx \
	$(FIXTURES_DIR)/edit_raw_copy.xlsx \
	$(FIXTURES_DIR)/edit_duplicate.zip \
	$(FIXTURES_DIR)/edit_duplicate_noop.zip
EDIT_FIXTURES := \
	$(FIXTURES_DIR)/edit_custom_target.xlsx \
	$(FIXTURES_DIR)/edit_normalized_target.xlsx
READER_FIXTURES := \
	$(FIXTURES_DIR)/types.xlsx \
	$(FIXTURES_DIR)/phase1.xlsx \
	$(FIXTURES_DIR)/phase3.xlsx \
	$(FIXTURES_DIR)/phase4.xlsx

PUBLIC_HEADERS := $(patsubst $(INC_DIR)/%,%, \
	$(wildcard $(INC_DIR)/libxlsx.h) \
	$(wildcard $(INC_DIR)/libxlsx/*.h))
HEADER_TEST_STAMPS := $(patsubst %.h,$(OBJ_DIR)/headers/%.ok,$(PUBLIC_HEADERS))

.PHONY: all test clean writer-test reader-test reader-fixtures integration-test edit-test edit-fixtures formula-test header-test

all: $(LIB_TEST)

test: writer-test reader-test integration-test edit-test formula-test header-test

writer-test: $(WRITER_TEST_BIN)
	$(Q)./$(WRITER_TEST_BIN)

reader-test: reader-fixtures $(READER_TEST_BINS)
	$(Q)set -e; for test_bin in $(READER_TEST_BINS); do printf '==> %s\n' "$$test_bin"; ./$$test_bin; done

integration-test: $(INTEGRATION_TEST_BINS)
	$(Q)set -e; for test_bin in $(INTEGRATION_TEST_BINS); do printf '==> %s\n' "$$test_bin"; ./$$test_bin; done

edit-test: edit-fixtures $(EDIT_TEST_BINS)
	$(Q)set -e; for test_bin in $(EDIT_TEST_BINS); do printf '==> %s\n' "$$test_bin"; ./$$test_bin; done

edit-fixtures: $(EDIT_FIXTURES)

formula-test: $(FORMULA_TEST_BINS)
	$(Q)set -e; for test_bin in $(FORMULA_TEST_BINS); do printf '==> %s\n' "$$test_bin"; ./$$test_bin; done

header-test: $(HEADER_TEST_STAMPS)

reader-fixtures: $(READER_FIXTURES)

$(LIB_TEST): $(LIB_OBJS)
	$(Q)$(AR) $(ARFLAGS) $@ $^

$(OBJ_DIR)/libxlsx $(OBJ_DIR)/minizip $(OBJ_DIR)/tmpfileplus $(OBJ_DIR)/md5 $(OBJ_DIR)/dtoa $(OBJ_DIR)/expat:
	$(Q)mkdir -p $@

$(OBJ_DIR)/libxlsx/%.o: $(SRC_DIR)/%.c | $(OBJ_DIR)/libxlsx
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/minizip/%.o: ../third_party/minizip/%.c | $(OBJ_DIR)/minizip
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/tmpfileplus/%.o: ../third_party/tmpfileplus/%.c | $(OBJ_DIR)/tmpfileplus
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/md5/%.o: ../third_party/md5/%.c | $(OBJ_DIR)/md5
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/dtoa/%.o: ../third_party/dtoa/%.c | $(OBJ_DIR)/dtoa
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/expat/%.o: $(EXPAT_DIR)/%.c | $(OBJ_DIR)/expat
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(OBJ_DIR)/writer/%.o: $(WRITER_TEST_DIR)/%.c
	$(Q)mkdir -p $(@D)
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $< -o $@

$(WRITER_TEST_BIN): $(WRITER_TEST_OBJS) $(LIB_TEST)
	$(Q)$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $^ $(LIBS_O)

$(READER_TEST_DIR)/%: $(READER_TEST_DIR)/%.c $(LIB_TEST) $(READER_TEST_DIR)/unity.h $(READER_TEST_DIR)/xlsx_test_paths.h
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIB_TEST) $(LIBS_O)

$(INTEGRATION_TEST_DIR)/%: $(INTEGRATION_TEST_DIR)/%.c $(LIB_TEST) $(READER_TEST_DIR)/unity.h
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIB_TEST) $(LIBS_O)

$(EDIT_TEST_DIR)/%: $(EDIT_TEST_DIR)/%.c $(LIB_TEST) $(READER_TEST_DIR)/unity.h
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIB_TEST) $(LIBS_O)

$(FORMULA_TEST_DIR)/%: $(FORMULA_TEST_DIR)/%.c $(LIB_TEST) $(READER_TEST_DIR)/unity.h
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIB_TEST) $(LIBS_O)

$(OBJ_DIR)/headers/%.ok: $(INC_DIR)/%.h
	$(Q)mkdir -p $(@D)
	$(Q)printf '#include "%s"\nint main(void) { return 0; }\n' "$*.h" > $(@:.ok=.c)
	$(Q)$(CC) $(CPPFLAGS) $(CFLAGS) -c $(@:.ok=.c) -o $(@:.ok=.o)
	$(Q)touch $@

$(FIXTURES_DIR)/types.xlsx: $(FIXTURES_DIR)/gen_types_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase1.xlsx: $(FIXTURES_DIR)/gen_phase1_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase3.xlsx: $(FIXTURES_DIR)/gen_phase3_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/phase4.xlsx: $(FIXTURES_DIR)/gen_phase4_xlsx.py
	$(Q)$(PYTHON) $< $@

$(FIXTURES_DIR)/edit_custom_target.xlsx: $(FIXTURES_DIR)/gen_edit_relationship_xlsx.py
	$(Q)$(PYTHON) $< $@ "worksheets/custom.xml"

$(FIXTURES_DIR)/edit_normalized_target.xlsx: $(FIXTURES_DIR)/gen_edit_relationship_xlsx.py
	$(Q)$(PYTHON) $< $@ "./worksheets/../worksheets/custom.xml"

clean:
	$(Q)rm -rf $(OBJ_DIR) $(LIB_TEST) $(WRITER_TEST_BIN) $(READER_TEST_BINS) $(READER_FIXTURES) \
		$(READER_TEST_DIR)/*.dSYM $(INTEGRATION_TEST_BINS) $(INTEGRATION_TEST_DIR)/*.dSYM \
		$(EDIT_TEST_BINS) $(EDIT_TEST_DIR)/*.dSYM \
		$(FORMULA_TEST_BINS) $(FORMULA_TEST_DIR)/*.dSYM \
		$(INTEGRATION_OUTPUTS) $(EDIT_OUTPUTS) $(EDIT_FIXTURES)
