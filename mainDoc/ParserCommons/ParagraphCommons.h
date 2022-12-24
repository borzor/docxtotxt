//
// Created by boris on 6/9/22.
//
#ifndef DOCXTOTXT_PARAGRAPHCOMMONS_H
#define DOCXTOTXT_PARAGRAPHCOMMONS_H

#include <queue>

namespace docxtotxt {
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
}
#endif //DOCXTOTXT_PARAGRAPHCOMMONS_H