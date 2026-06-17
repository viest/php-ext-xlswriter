/*****************************************************************************
 * comment - A library for creating Excel XLSX comment files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/comment.h"
#include "lxlsx/utility.h"

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
