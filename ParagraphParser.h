//
// Created by boris on 2/20/22.
//
#include "tinyxml2.h"
#include "tinyxml.h"
#include <zip.h>
#include "iostream"
#include <map>
#include <vector>
#include <fstream>

#ifndef DOCXTOTXT_PARAGRAPHBUFFER_H
#define DOCXTOTXT_PARAGRAPHBUFFER_H

namespace paragraph {
    using namespace std;
    using namespace tinyxml2;
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
    enum paragraphJustify{
        left,
        right,
        center,
        both,
        distribute
    };
    static map<string, int> paragraphProperties = {{"w:framePr",       framePr},
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
    static map<string, int> textProperties = {{"w:br",            br},
                                              {"w:cr",            cr},
                                              {"w:drawing",       drawing},
                                              {"w:noBreakHyphen", noBreakHyphen},
                                              {"w:rPr",           rPr},
                                              {"w:softHyphen",    softHyphen},
                                              {"w:sym",           sym},
                                              {"w:t",             t},
                                              {"w:tab",           tab},};

    class ParagraphParser {
    private:
        vector<string> paragraphBuffer;
        paragraphJustify justify;

    public:
        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void addText(const string& text);

        void setJustify(const string& justify);

        void writeResult(ofstream &outfile, bool toFile);
    };

}

#endif //DOCXTOTXT_PARAGRAPHBUFFER_H
