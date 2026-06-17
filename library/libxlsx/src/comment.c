/*****************************************************************************
 * comment - A library for creating Excel XLSX comment files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/comment.h"
#include "libxlsx/utility.h"

/*
 * Forward declarations.
 */

STATIC int _author_id_cmp(lxlsx_author_id *tuple1, lxlsx_author_id *tuple2);

#ifndef __clang_analyzer__
LXLSX_RB_GENERATE_AUTHOR_IDS(lxlsx_author_ids, lxlsx_author_id,
                           tree_pointers, _author_id_cmp);
#endif

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Comparator for the author ids.
 */
STATIC int
_author_id_cmp(lxlsx_author_id *author_id1, lxlsx_author_id *author_id2)
{
    return strcmp(author_id1->author, author_id2->author);
}

/*
 * Check if an author already existing in the author/id table.
 */
STATIC uint8_t
_check_author(lxlsx_comment *self, char *author)
{
    lxlsx_author_id tmp_author_id;
    lxlsx_author_id *existing_author = NULL;

    if (!author)
        return LXLSX_TRUE;

    tmp_author_id.author = author;
    existing_author = RB_FIND(lxlsx_author_ids,
                              self->author_ids, &tmp_author_id);

    if (existing_author)
        return LXLSX_TRUE;
    else
        return LXLSX_FALSE;
}

/*
 * Get the index used for an author name.
 */
STATIC uint32_t
_get_author_index(lxlsx_comment *self, char *author)
{
    lxlsx_author_id tmp_author_id;
    lxlsx_author_id *existing_author = NULL;
    lxlsx_author_id *new_author_id = NULL;

    if (!author)
        return 0;

    tmp_author_id.author = author;
    existing_author = RB_FIND(lxlsx_author_ids,
                              self->author_ids, &tmp_author_id);

    if (existing_author) {
        return existing_author->id;
    }
    else {
        new_author_id = calloc(1, sizeof(lxlsx_author_id));

        if (new_author_id) {
            new_author_id->id = self->author_id;
            new_author_id->author = lxlsx_strdup(author);
            self->author_id++;

            RB_INSERT(lxlsx_author_ids, self->author_ids, new_author_id);

            return new_author_id->id;
        }
        else {
            return 0;
        }
    }
}

/*
 * Create a new comment object.
 */
lxlsx_comment *
lxlsx_comment_new(void)
{
    lxlsx_comment *comment = calloc(1, sizeof(lxlsx_comment));
    GOTO_LABEL_ON_MEM_ERROR(comment, mem_error);

    comment->author_ids = calloc(1, sizeof(struct lxlsx_author_ids));
    GOTO_LABEL_ON_MEM_ERROR(comment->author_ids, mem_error);
    RB_INIT(comment->author_ids);

    return comment;

mem_error:
    lxlsx_comment_free(comment);
    return NULL;
}

/*
 * Free a comment object.
 */
void
lxlsx_comment_free(lxlsx_comment *comment)
{
    struct lxlsx_author_id *author_id;
    struct lxlsx_author_id *next_author_id;

    if (!comment)
        return;

    if (comment->author_ids) {
        for (author_id =
             RB_MIN(lxlsx_author_ids, comment->author_ids);
             author_id; author_id = next_author_id) {

            next_author_id =
                RB_NEXT(lxlsx_author_ids, worksheet->author_id, author_id);
            RB_REMOVE(lxlsx_author_ids, comment->author_ids, author_id);
            free(author_id->author);
            free(author_id);
        }

        free(comment->author_ids);
    }

    free(comment);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
STATIC void
_comment_xml_declaration(lxlsx_comment *self)
{
    lxlsx_xml_declaration(self->file);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <t> element.
 */
STATIC void
_comment_write_text_t(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    lxlsx_xml_data_element(self->file, "t", comment_obj->text, NULL);
}

/*
 * Write the <family> element.
 */
STATIC void
_comment_write_family(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("val", comment_obj->font_family);

    lxlsx_xml_empty_tag(self->file, "family", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <rFont> element.
 */
STATIC void
_comment_write_r_font(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char font_name[LXLSX_ATTR_32];

    if (comment_obj->font_name)
        lxlsx_snprintf(font_name, LXLSX_ATTR_32, "%s", comment_obj->font_name);
    else
        lxlsx_snprintf(font_name, LXLSX_ATTR_32, "Tahoma");

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("val", font_name);

    lxlsx_xml_empty_tag(self->file, "rFont", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <color> element.
 */
STATIC void
_comment_write_color(lxlsx_comment *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char indexed[] = "81";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("indexed", indexed);

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sz> element.
 */
STATIC void
_comment_write_sz(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_DBL("val", comment_obj->font_size);

    lxlsx_xml_empty_tag(self->file, "sz", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <rPr> element.
 */
STATIC void
_comment_write_r_pr(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    lxlsx_xml_start_tag(self->file, "rPr", NULL);

    /* Write the sz element. */
    _comment_write_sz(self, comment_obj);

    /* Write the color element. */
    _comment_write_color(self);

    /* Write the rFont element. */
    _comment_write_r_font(self, comment_obj);

    /* Write the family element. */
    _comment_write_family(self, comment_obj);

    lxlsx_xml_end_tag(self->file, "rPr");
}

/*
 * Write the <r> element.
 */
STATIC void
_comment_write_r(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    lxlsx_xml_start_tag(self->file, "r", NULL);

    /* Write the rPr element. */
    _comment_write_r_pr(self, comment_obj);

    /* Write the t element. */
    _comment_write_text_t(self, comment_obj);

    lxlsx_xml_end_tag(self->file, "r");
}

/*
 * Write the <text> element.
 */
STATIC void
_comment_write_text(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    lxlsx_xml_start_tag(self->file, "text", NULL);

    /* Write the r element. */
    _comment_write_r(self, comment_obj);

    lxlsx_xml_end_tag(self->file, "text");
}

/*
 * Write the <comment> element.
 */
STATIC void
_comment_write_comment(lxlsx_comment *self, lxlsx_vml_obj *comment_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char ref[LXLSX_MAX_CELL_NAME_LENGTH];

    lxlsx_rowcol_to_cell(ref, comment_obj->row, comment_obj->col);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", ref);
    LXLSX_PUSH_ATTRIBUTES_INT("authorId", comment_obj->author_id);

    lxlsx_xml_start_tag(self->file, "comment", &attributes);

    /* Write the text element. */
    _comment_write_text(self, comment_obj);

    lxlsx_xml_end_tag(self->file, "comment");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <commentList> element.
 */
STATIC void
_comment_write_comment_list(lxlsx_comment *self)
{
    lxlsx_vml_obj *comment_obj;

    lxlsx_xml_start_tag(self->file, "commentList", NULL);

    STAILQ_FOREACH(comment_obj, self->comment_objs, list_pointers) {
        /* Write the comment element. */
        _comment_write_comment(self, comment_obj);
    }

    lxlsx_xml_end_tag(self->file, "commentList");

}

/*
 * Write the <author> element.
 */
STATIC void
_comment_write_author(lxlsx_comment *self, char *author)
{
    lxlsx_xml_data_element(self->file, "author", author, NULL);
}

/*
 * Write the <authors> element.
 */
STATIC void
_comment_write_authors(lxlsx_comment *self)
{
    lxlsx_vml_obj *comment_obj;
    char *author;

    lxlsx_xml_start_tag(self->file, "authors", NULL);

    /* Set the default author (from lxlsx_worksheet_set_comments_author()). */
    if (self->comment_author) {
        _get_author_index(self, self->comment_author);
        _comment_write_author(self, self->comment_author);
    }
    else {
        _get_author_index(self, "");
        _comment_write_author(self, "");
    }

    STAILQ_FOREACH(comment_obj, self->comment_objs, list_pointers) {
        author = comment_obj->author;

        if (author) {

            if (!_check_author(self, author))
                _comment_write_author(self, author);

            comment_obj->author_id = _get_author_index(self, author);
        }
    }

    lxlsx_xml_end_tag(self->file, "authors");
}

/*
 * Write the <comments> element.
 */
STATIC void
_comment_write_comments(lxlsx_comment *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);

    lxlsx_xml_start_tag(self->file, "comments", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_comment_assemble_xml_file(lxlsx_comment *self)
{
    /* Write the XML declaration. */
    _comment_xml_declaration(self);

    /* Write the comments element. */
    _comment_write_comments(self);

    /* Write the authors element. */
    _comment_write_authors(self);

    /* Write the commentList element. */
    _comment_write_comment_list(self);

    lxlsx_xml_end_tag(self->file, "comments");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/


/****************************************************************************
 *
 * XLSX comment read support.
 *
 ****************************************************************************/

/*
 * Comments reader.
 *
 * Iterates legacy comments (xl/comments*.xml + xl/drawings/vmlDrawing*.vml
 * for visibility) and threaded comments (xl/threadedComments/threadedComment*.xml)
 * for a given worksheet. The sheet's _rels file is the entry point — we
 * resolve the relationship targets there.
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "xlsx_private.h"
#include "xlsx_util.h"

/* --- comments XML parse ------------------------------------------------- */

typedef struct {
    char *guid;        /* author GUID for threaded; placeholder index for legacy */
    char *display;
} c_author;

typedef struct {
    size_t row;
    size_t col;
    char  *text;        /* concatenated */
    char  *author;
    char  *guid;        /* for threaded */
    char  *parent_guid;
    int    visible;
    int    threaded;
} c_record;

typedef struct {
    /* Authors table (legacy & threaded share via parallel arrays). */
    c_author *authors;
    size_t    authors_count;
    size_t    authors_cap;
    int       legacy_author_idx;    /* used when collecting authorId-indexed legacy authors */

    /* Records pool (both legacy and threaded mixed in). */
    c_record *records;
    size_t    records_count;
    size_t    records_cap;

    /* SAX state. */
    int   in_authors;
    int   in_author;
    int   in_commentList;
    int   in_comment;
    int   in_text;          /* text/<r>/<t> chain */
    int   in_run;
    int   in_run_t;
    char *cur_text;
    size_t cur_text_len;
    size_t cur_text_cap;
    char *author_text;
    size_t author_text_len;
    size_t author_text_cap;
    /* threaded specifics */
    int   in_threaded_root;
    int   in_threaded_text;
    char *th_pending_id;
    char *th_pending_parent;
    char *th_pending_ref;
    char *th_pending_author_id;
} c_ctx;

/* --- legacy comments SAX (xl/comments*.xml) ----------------------------- */

static c_record *records_push(c_ctx *c)
{
    if (c->records_count >= c->records_cap) {
        size_t nc = c->records_cap ? c->records_cap * 2 : 8;
        c_record *nb = (c_record *)realloc(c->records, nc * sizeof(*nb));
        if (!nb) return NULL;
        c->records = nb; c->records_cap = nc;
    }
    {
        c_record *r = &c->records[c->records_count++];
        memset(r, 0, sizeof(*r));
        return r;
    }
}

static void on_legacy_start(void *ud, const char *name, const char **attrs)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxlsx_reader_xml_name_eq(name, "authors")) { c->in_authors = 1; c->legacy_author_idx = 0; return; }
    if (lxlsx_reader_xml_name_eq(name, "commentList")) { c->in_commentList = 1; return; }
    if (c->in_authors && lxlsx_reader_xml_name_eq(name, "author")) {
        c->in_author = 1;
        if (c->authors_count >= c->authors_cap) {
            size_t nc = c->authors_cap ? c->authors_cap * 2 : 4;
            c_author *nb = (c_author *)realloc(c->authors, nc * sizeof(*nb));
            if (nb) { c->authors = nb; c->authors_cap = nc; }
        }
        if (c->authors_count < c->authors_cap) {
            char buf[24];
            snprintf(buf, sizeof(buf), "%d", c->legacy_author_idx);
            c->authors[c->authors_count].guid    = strdup(buf);
            c->authors[c->authors_count].display = NULL;
            c->authors_count++;
            c->legacy_author_idx++;
        }
        lxlsx_reader_buf_reset(c->author_text, &c->author_text_len);
        return;
    }
    if (c->in_commentList && lxlsx_reader_xml_name_eq(name, "comment")) {
        const char *ref      = lxlsx_reader_xml_attr(attrs, "ref");
        const char *authorId = lxlsx_reader_xml_attr(attrs, "authorId");
        c_record   *r        = records_push(c);
        if (r) {
            lxlsx_reader_parse_a1_ref(ref, &r->row, &r->col);
            if (authorId) {
                int idx = (int)strtol(authorId, NULL, 10);
                if (idx >= 0 && (size_t)idx < c->authors_count
                              && c->authors[idx].display) {
                    r->author = strdup(c->authors[idx].display);
                }
            }
            c->in_comment = 1;
            lxlsx_reader_buf_reset(c->cur_text, &c->cur_text_len);
        }
        return;
    }
    if (c->in_comment) {
        if (lxlsx_reader_xml_name_eq(name, "text"))    c->in_text = 1;
        else if (c->in_text && lxlsx_reader_xml_name_eq(name, "r"))   c->in_run = 1;
        else if (c->in_text && lxlsx_reader_xml_name_eq(name, "t"))   c->in_run_t = 1;
        else if (c->in_run && lxlsx_reader_xml_name_eq(name, "t"))    c->in_run_t = 1;
    }
}

static void on_legacy_end(void *ud, const char *name)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxlsx_reader_xml_name_eq(name, "authors"))     c->in_authors = 0;
    if (lxlsx_reader_xml_name_eq(name, "commentList")) c->in_commentList = 0;
    if (c->in_author && lxlsx_reader_xml_name_eq(name, "author")) {
        if (c->authors_count > 0) {
            free(c->authors[c->authors_count - 1].display);
            c->authors[c->authors_count - 1].display =
                c->author_text_len > 0 ? strdup(c->author_text) : NULL;
        }
        c->in_author = 0;
    }
    if (c->in_run_t && lxlsx_reader_xml_name_eq(name, "t")) c->in_run_t = 0;
    if (c->in_run   && lxlsx_reader_xml_name_eq(name, "r")) c->in_run = 0;
    if (c->in_text  && lxlsx_reader_xml_name_eq(name, "text")) c->in_text = 0;
    if (c->in_comment && lxlsx_reader_xml_name_eq(name, "comment")) {
        if (c->records_count > 0 && c->cur_text_len > 0) {
            c_record *r = &c->records[c->records_count - 1];
            r->text = strdup(c->cur_text);
        }
        c->in_comment = 0;
    }
}

static void on_legacy_text(void *ud, const char *text, int len)
{
    c_ctx *c = (c_ctx *)ud;
    if (len <= 0) return;
    if (c->in_author && !c->in_run_t)
        lxlsx_reader_buf_append(&c->author_text, &c->author_text_len,
                                &c->author_text_cap, text, (size_t)len);
    if (c->in_comment && c->in_run_t)
        lxlsx_reader_buf_append(&c->cur_text, &c->cur_text_len,
                                &c->cur_text_cap, text, (size_t)len);
}

/* --- threadedComments parse --------------------------------------------- */

static void on_thread_start(void *ud, const char *name, const char **attrs)
{
    c_ctx *c = (c_ctx *)ud;
    /* Microsoft's threadedComments root is conventionally <ThreadedComments>
     * (capital T), but accept the lowercase form too. */
    if (lxlsx_reader_xml_name_eq(name, "ThreadedComments") ||
        lxlsx_reader_xml_name_eq(name, "threadedComments")) { c->in_threaded_root = 1; return; }
    if (!c->in_threaded_root) return;
    if (lxlsx_reader_xml_name_eq(name, "threadedComment")) {
        const char *ref      = lxlsx_reader_xml_attr(attrs, "ref");
        const char *id       = lxlsx_reader_xml_attr(attrs, "id");
        const char *parent   = lxlsx_reader_xml_attr(attrs, "parentId");
        const char *authorId = lxlsx_reader_xml_attr(attrs, "personId");
        free(c->th_pending_id);        c->th_pending_id        = id       ? strdup(id) : NULL;
        free(c->th_pending_parent);    c->th_pending_parent    = parent   ? strdup(parent) : NULL;
        free(c->th_pending_ref);       c->th_pending_ref       = ref      ? strdup(ref) : NULL;
        free(c->th_pending_author_id); c->th_pending_author_id = authorId ? strdup(authorId) : NULL;
        lxlsx_reader_buf_reset(c->cur_text, &c->cur_text_len);
        c->in_threaded_text = 0;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "text")) c->in_threaded_text = 1;
}

static void on_thread_end(void *ud, const char *name)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxlsx_reader_xml_name_eq(name, "text")) c->in_threaded_text = 0;
    if (lxlsx_reader_xml_name_eq(name, "ThreadedComments") ||
        lxlsx_reader_xml_name_eq(name, "threadedComments")) c->in_threaded_root = 0;
    if (lxlsx_reader_xml_name_eq(name, "threadedComment")) {
        c_record *r = records_push(c);
        if (r) {
            r->threaded = 1;
            r->visible  = 0;
            lxlsx_reader_parse_a1_ref(c->th_pending_ref, &r->row, &r->col);
            r->guid          = c->th_pending_id;        c->th_pending_id = NULL;
            r->parent_guid   = c->th_pending_parent;    c->th_pending_parent = NULL;
            r->author        = c->th_pending_author_id; c->th_pending_author_id = NULL;
            r->text          = c->cur_text_len > 0 ? strdup(c->cur_text) : NULL;
        }
    }
}

static void on_thread_text(void *ud, const char *text, int len)
{
    c_ctx *c = (c_ctx *)ud;
    if (c->in_threaded_text && len > 0)
        lxlsx_reader_buf_append(&c->cur_text, &c->cur_text_len,
                                &c->cur_text_cap, text, (size_t)len);
}

/* --- vmlDrawing visibility extraction (legacy) -------------------------- */

typedef struct {
    /* row/col -> visibility map. We just collect (row,col,visible). */
    struct { size_t row; size_t col; int visible; } *items;
    size_t count;
    size_t cap;
} vml_visibility;

static void vml_push(vml_visibility *v, size_t row, size_t col, int visible)
{
    if (v->count >= v->cap) {
        size_t nc = v->cap ? v->cap * 2 : 8;
        void *nb = realloc(v->items, nc * sizeof(*v->items));
        if (!nb) return;
        v->items = nb; v->cap = nc;
    }
    v->items[v->count].row     = row;
    v->items[v->count].col     = col;
    v->items[v->count].visible = visible;
    v->count++;
}

static int vml_lookup(const vml_visibility *v, size_t row, size_t col)
{
    size_t i;
    if (!v) return 0;
    for (i = 0; i < v->count; i++) {
        if (v->items[i].row == row && v->items[i].col == col)
            return v->items[i].visible;
    }
    return 0;
}

typedef struct {
    vml_visibility *v;
    int             in_shape;
    int             in_clientdata;
    int             in_anchor;
    int             in_visible;
    int             in_row;
    int             in_column;
    int             pending_visible;
    long            pending_col;       /* 0-based VML uses col, row */
    long            pending_row;
    char           *anchor_text;
    size_t          anchor_text_len;
    size_t          anchor_text_cap;
    char           *small_buf;
    size_t          small_buf_len;
    size_t          small_buf_cap;
} vml_ctx;

/* The VML format (legacy Excel comments) is a subset of legacy MS Office
 * shapes: each <v:shape> with an <x:ClientData ObjectType="Note"> child
 * that contains <x:Visible/>, <x:Row>R</x:Row>, <x:Column>C</x:Column>.
 * Tag prefixes vary; lxlsx_reader_xml_name_eq strips namespaces. */
static void vml_on_start(void *ud, const char *name, const char **attrs)
{
    vml_ctx *c = (vml_ctx *)ud;
    (void)attrs;
    if (lxlsx_reader_xml_name_eq(name, "shape")) {
        c->in_shape = 1; c->pending_visible = 0;
        c->pending_row = c->pending_col = -1;
    }
    if (c->in_shape && lxlsx_reader_xml_name_eq(name, "ClientData")) c->in_clientdata = 1;
    if (c->in_clientdata && lxlsx_reader_xml_name_eq(name, "Visible")) c->pending_visible = 1;
    if (c->in_clientdata && lxlsx_reader_xml_name_eq(name, "Row"))     { c->in_row = 1; c->small_buf_len = 0; if (c->small_buf) c->small_buf[0] = 0; }
    if (c->in_clientdata && lxlsx_reader_xml_name_eq(name, "Column"))  { c->in_column = 1; c->small_buf_len = 0; if (c->small_buf) c->small_buf[0] = 0; }
}

static void vml_on_text(void *ud, const char *text, int len)
{
    vml_ctx *c = (vml_ctx *)ud;
    if (len <= 0) return;
    if (c->in_row || c->in_column) {
        lxlsx_reader_buf_append(&c->small_buf, &c->small_buf_len,
                                &c->small_buf_cap, text, (size_t)len);
    }
}

static void vml_on_end(void *ud, const char *name)
{
    vml_ctx *c = (vml_ctx *)ud;
    if (c->in_row && lxlsx_reader_xml_name_eq(name, "Row")) {
        c->pending_row = strtol(c->small_buf ? c->small_buf : "-1", NULL, 10);
        c->in_row = 0;
    }
    if (c->in_column && lxlsx_reader_xml_name_eq(name, "Column")) {
        c->pending_col = strtol(c->small_buf ? c->small_buf : "-1", NULL, 10);
        c->in_column = 0;
    }
    if (c->in_clientdata && lxlsx_reader_xml_name_eq(name, "ClientData")) {
        if (c->pending_row >= 0 && c->pending_col >= 0) {
            vml_push(c->v, (size_t)(c->pending_row + 1),
                     (size_t)(c->pending_col + 1), c->pending_visible);
        }
        c->in_clientdata = 0;
    }
    if (c->in_shape && lxlsx_reader_xml_name_eq(name, "shape")) {
        c->in_shape = 0;
    }
}

static lxlsx_reader_error parse_vml(lxlsx_reader_zip *zip, const char *path, vml_visibility *v)
{
    lxlsx_reader_zip_file *zf = lxlsx_reader_zip_open_entry(zip, path);
    lxlsx_reader_xml_pump *pump;
    vml_ctx       ctx;
    lxlsx_reader_error     rc;
    if (!zf) return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) { lxlsx_reader_zip_close_entry(zf); return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED; }
    memset(&ctx, 0, sizeof(ctx));
    ctx.v = v;
    lxlsx_reader_xml_pump_set_handlers(pump, vml_on_start, vml_on_end, vml_on_text, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);
    free(ctx.small_buf);
    free(ctx.anchor_text);
    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    return rc;
}

/* --- public entry ------------------------------------------------------- */

static void free_records(c_record *records, size_t n)
{
    size_t i;
    for (i = 0; i < n; i++) {
        free(records[i].text);
        free(records[i].author);
        free(records[i].guid);
        free(records[i].parent_guid);
    }
    free(records);
}

static void free_authors(c_author *authors, size_t n)
{
    size_t i;
    for (i = 0; i < n; i++) {
        free(authors[i].guid);
        free(authors[i].display);
    }
    free(authors);
}

static lxlsx_reader_error scan_one_xml(lxlsx_reader_zip *zip, const char *path,
                              c_ctx *ctx, int threaded)
{
    lxlsx_reader_zip_file *zf = lxlsx_reader_zip_open_entry(zip, path);
    lxlsx_reader_xml_pump *pump;
    lxlsx_reader_error     rc;
    if (!zf) return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) { lxlsx_reader_zip_close_entry(zf); return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED; }
    if (threaded) {
        lxlsx_reader_xml_pump_set_handlers(pump, on_thread_start, on_thread_end, on_thread_text, ctx);
    } else {
        lxlsx_reader_xml_pump_set_handlers(pump, on_legacy_start, on_legacy_end, on_legacy_text, ctx);
    }
    rc = lxlsx_reader_xml_pump_run(pump);
    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    return rc;
}

lxlsx_reader_error lxlsx_reader_worksheet_iterate_comments(lxlsx_reader_worksheet *ws,
                                         lxlsx_reader_comment_cb cb, void *userdata)
{
    lxlsx_reader_rel_map rels;
    char           *comments_target = NULL;
    char           *thread_target   = NULL;
    char           *vml_target      = NULL;
    char           *comments_path   = NULL;
    char           *thread_path     = NULL;
    char           *vml_path        = NULL;
    c_ctx           ctx;
    vml_visibility  vml;
    lxlsx_reader_error       rc;
    size_t          i;

    if (!ws || !ws->wb || !cb) return LXLSX_READER_ERROR_NULL_PARAMETER;

    /* Collect ALL relationships once. */
    rc = lxlsx_reader_load_rels(ws->wb->zip, ws->target_path, &rels, 0);
    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_rel_map_free(&rels);
        return rc;
    }

    /* Walk the rels map and bucket. */
    for (i = 0; i < rels.count; i++) {
        const char *t = rels.items[i].target;
        if (!t) continue;
        if (strstr(t, "comments") && !strstr(t, "threaded") && !comments_target) {
            comments_target = lxlsx_reader_strdup(t);
        }
        if (strstr(t, "threadedComment") && !thread_target) {
            thread_target = lxlsx_reader_strdup(t);
        }
        if (strstr(t, "vmlDrawing") && !vml_target) {
            vml_target = lxlsx_reader_strdup(t);
        }
    }

    lxlsx_reader_rel_map_free(&rels);

    if (!comments_target && !thread_target) {
        free(vml_target);
        return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;
    }

    comments_path = comments_target
        ? lxlsx_reader_zip_resolve_path(ws->target_path, comments_target) : NULL;
    thread_path = thread_target
        ? lxlsx_reader_zip_resolve_path(ws->target_path, thread_target) : NULL;
    vml_path = vml_target
        ? lxlsx_reader_zip_resolve_path(ws->target_path, vml_target) : NULL;
    free(comments_target);
    free(thread_target);
    free(vml_target);

    memset(&ctx, 0, sizeof(ctx));
    memset(&vml, 0, sizeof(vml));

    if (comments_path) (void)scan_one_xml(ws->wb->zip, comments_path, &ctx, 0);
    if (thread_path)   (void)scan_one_xml(ws->wb->zip, thread_path,   &ctx, 1);
    if (vml_path)      (void)parse_vml(ws->wb->zip, vml_path, &vml);

    free(comments_path); free(thread_path); free(vml_path);

    /* Deliver. */
    {
        int stop = 0;
        for (i = 0; i < ctx.records_count && !stop; i++) {
            lxlsx_reader_comment_info info;
            memset(&info, 0, sizeof(info));
            info.row       = ctx.records[i].row;
            info.col       = ctx.records[i].col;
            info.text      = ctx.records[i].text;
            info.author    = ctx.records[i].author;
            info.threaded  = ctx.records[i].threaded;
            info.parent_id = ctx.records[i].parent_guid;
            info.visible   = ctx.records[i].threaded ? 0 :
                             vml_lookup(&vml, info.row, info.col);
            if (cb(&info, userdata) != 0) stop = 1;
        }
    }

    free_records(ctx.records, ctx.records_count);
    free_authors(ctx.authors, ctx.authors_count);
    free(ctx.cur_text);
    free(ctx.author_text);
    free(ctx.th_pending_id);
    free(ctx.th_pending_parent);
    free(ctx.th_pending_ref);
    free(ctx.th_pending_author_id);
    free(vml.items);
    return LXLSX_READER_NO_ERROR;
}
