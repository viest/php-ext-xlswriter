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

STATIC int _author_id_cmp(lxw_author_id *tuple1, lxw_author_id *tuple2);

#ifndef __clang_analyzer__
LXW_RB_GENERATE_AUTHOR_IDS(lxw_author_ids, lxw_author_id,
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
_author_id_cmp(lxw_author_id *author_id1, lxw_author_id *author_id2)
{
    return strcmp(author_id1->author, author_id2->author);
}

/*
 * Check if an author already existing in the author/id table.
 */
STATIC uint8_t
_check_author(lxw_comment *self, char *author)
{
    lxw_author_id tmp_author_id;
    lxw_author_id *existing_author = NULL;

    if (!author)
        return LXW_TRUE;

    tmp_author_id.author = author;
    existing_author = RB_FIND(lxw_author_ids,
                              self->author_ids, &tmp_author_id);

    if (existing_author)
        return LXW_TRUE;
    else
        return LXW_FALSE;
}

/*
 * Get the index used for an author name.
 */
STATIC uint32_t
_get_author_index(lxw_comment *self, char *author)
{
    lxw_author_id tmp_author_id;
    lxw_author_id *existing_author = NULL;
    lxw_author_id *new_author_id = NULL;

    if (!author)
        return 0;

    tmp_author_id.author = author;
    existing_author = RB_FIND(lxw_author_ids,
                              self->author_ids, &tmp_author_id);

    if (existing_author) {
        return existing_author->id;
    }
    else {
        new_author_id = calloc(1, sizeof(lxw_author_id));

        if (new_author_id) {
            new_author_id->id = self->author_id;
            new_author_id->author = lxw_strdup(author);
            self->author_id++;

            RB_INSERT(lxw_author_ids, self->author_ids, new_author_id);

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
lxw_comment *
lxw_comment_new(void)
{
    lxw_comment *comment = calloc(1, sizeof(lxw_comment));
    GOTO_LABEL_ON_MEM_ERROR(comment, mem_error);

    comment->author_ids = calloc(1, sizeof(struct lxw_author_ids));
    GOTO_LABEL_ON_MEM_ERROR(comment->author_ids, mem_error);
    RB_INIT(comment->author_ids);

    return comment;

mem_error:
    lxw_comment_free(comment);
    return NULL;
}

/*
 * Free a comment object.
 */
void
lxw_comment_free(lxw_comment *comment)
{
    struct lxw_author_id *author_id;
    struct lxw_author_id *next_author_id;

    if (!comment)
        return;

    if (comment->author_ids) {
        for (author_id =
             RB_MIN(lxw_author_ids, comment->author_ids);
             author_id; author_id = next_author_id) {

            next_author_id =
                RB_NEXT(lxw_author_ids, worksheet->author_id, author_id);
            RB_REMOVE(lxw_author_ids, comment->author_ids, author_id);
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
_comment_xml_declaration(lxw_comment *self)
{
    lxw_xml_declaration(self->file);
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
_comment_write_text_t(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    lxw_xml_data_element(self->file, "t", comment_obj->text, NULL);
}

/*
 * Write the <family> element.
 */
STATIC void
_comment_write_family(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", comment_obj->font_family);

    lxw_xml_empty_tag(self->file, "family", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <rFont> element.
 */
STATIC void
_comment_write_r_font(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char font_name[LXW_ATTR_32];

    if (comment_obj->font_name)
        lxw_snprintf(font_name, LXW_ATTR_32, "%s", comment_obj->font_name);
    else
        lxw_snprintf(font_name, LXW_ATTR_32, "Tahoma");

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", font_name);

    lxw_xml_empty_tag(self->file, "rFont", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <color> element.
 */
STATIC void
_comment_write_color(lxw_comment *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char indexed[] = "81";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("indexed", indexed);

    lxw_xml_empty_tag(self->file, "color", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <sz> element.
 */
STATIC void
_comment_write_sz(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", comment_obj->font_size);

    lxw_xml_empty_tag(self->file, "sz", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <rPr> element.
 */
STATIC void
_comment_write_r_pr(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    lxw_xml_start_tag(self->file, "rPr", NULL);

    /* Write the sz element. */
    _comment_write_sz(self, comment_obj);

    /* Write the color element. */
    _comment_write_color(self);

    /* Write the rFont element. */
    _comment_write_r_font(self, comment_obj);

    /* Write the family element. */
    _comment_write_family(self, comment_obj);

    lxw_xml_end_tag(self->file, "rPr");
}

/*
 * Write the <r> element.
 */
STATIC void
_comment_write_r(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    lxw_xml_start_tag(self->file, "r", NULL);

    /* Write the rPr element. */
    _comment_write_r_pr(self, comment_obj);

    /* Write the t element. */
    _comment_write_text_t(self, comment_obj);

    lxw_xml_end_tag(self->file, "r");
}

/*
 * Write the <text> element.
 */
STATIC void
_comment_write_text(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    lxw_xml_start_tag(self->file, "text", NULL);

    /* Write the r element. */
    _comment_write_r(self, comment_obj);

    lxw_xml_end_tag(self->file, "text");
}

/*
 * Write the <comment> element.
 */
STATIC void
_comment_write_comment(lxw_comment *self, lxw_vml_obj *comment_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char ref[LXW_MAX_CELL_NAME_LENGTH];

    lxw_rowcol_to_cell(ref, comment_obj->row, comment_obj->col);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", ref);
    LXW_PUSH_ATTRIBUTES_INT("authorId", comment_obj->author_id);

    lxw_xml_start_tag(self->file, "comment", &attributes);

    /* Write the text element. */
    _comment_write_text(self, comment_obj);

    lxw_xml_end_tag(self->file, "comment");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <commentList> element.
 */
STATIC void
_comment_write_comment_list(lxw_comment *self)
{
    lxw_vml_obj *comment_obj;

    lxw_xml_start_tag(self->file, "commentList", NULL);

    STAILQ_FOREACH(comment_obj, self->comment_objs, list_pointers) {
        /* Write the comment element. */
        _comment_write_comment(self, comment_obj);
    }

    lxw_xml_end_tag(self->file, "commentList");

}

/*
 * Write the <author> element.
 */
STATIC void
_comment_write_author(lxw_comment *self, char *author)
{
    lxw_xml_data_element(self->file, "author", author, NULL);
}

/*
 * Write the <authors> element.
 */
STATIC void
_comment_write_authors(lxw_comment *self)
{
    lxw_vml_obj *comment_obj;
    char *author;

    lxw_xml_start_tag(self->file, "authors", NULL);

    /* Set the default author (from worksheet_set_comments_author()). */
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

    lxw_xml_end_tag(self->file, "authors");
}

/*
 * Write the <comments> element.
 */
STATIC void
_comment_write_comments(lxw_comment *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);

    lxw_xml_start_tag(self->file, "comments", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxw_comment_assemble_xml_file(lxw_comment *self)
{
    /* Write the XML declaration. */
    _comment_xml_declaration(self);

    /* Write the comments element. */
    _comment_write_comments(self);

    /* Write the authors element. */
    _comment_write_authors(self);

    /* Write the commentList element. */
    _comment_write_comment_list(self);

    lxw_xml_end_tag(self->file, "comments");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
