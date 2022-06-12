//
// Created by boris on 6/9/22.
//
#ifndef DOCXTOTXT_PARAGRAPHCOMMONS_H
#define DOCXTOTXT_PARAGRAPHCOMMONS_H

enum paragraphProperty {
    framePr,
    ind,
    jc,
    keepLines,
    keepNext,
    numPr,
    outlineLvl,
    pBdr,
    pStyle,
    rPr,
    sectPr,
    shd,
    spacing,
    tabs,
    textAlignment
};
enum textProperty {
    br,
    cr,
    drawing,
    noBreakHyphen,
    softHyphen,
    sym,
    t,
    tab
};
enum paragraphJustify {
    left,
    right,
    center,
    both,
    distribute
};
typedef struct{
    size_t left;
    size_t right;
    size_t hanging;
    size_t firstLine;
} indentation;

typedef struct {
    paragraphJustify justify;
    indentation ind;
} paragraphSettings;

typedef struct {
} textSettings;

static std::map<std::string, int> paragraphProperties = {{"w:framePr",       framePr},
                                                         {"w:ind",           ind},
                                                         {"w:jc",            jc},
                                                         {"w:keepLines",     keepLines},
                                                         {"w:keepNext",      keepNext},
                                                         {"w:numPr",         numPr},
                                                         {"w:outlineLvl",    outlineLvl},
                                                         {"w:pBdr",          pBdr},
                                                         {"w:pStyle",        pStyle},
                                                         {"w:rPr",           rPr},
                                                         {"w:sectPr",        sectPr},
                                                         {"w:shd",           shd},
                                                         {"w:spacing",       spacing},
                                                         {"w:tabs",          tabs},
                                                         {"w:textAlignment", textAlignment}};
static std::map<std::string, int> textProperties = {{"w:br",            br},
                                                    {"w:cr",            cr},
                                                    {"w:drawing",       drawing},
                                                    {"w:noBreakHyphen", noBreakHyphen},
                                                    {"w:rPr",           rPr},
                                                    {"w:softHyphen",    softHyphen},
                                                    {"w:sym",           sym},
                                                    {"w:t",             t},
                                                    {"w:tab",           tab},};
#endif //DOCXTOTXT_PARAGRAPHCOMMONS_H