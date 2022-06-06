//
// Created by boris on 2/20/22.
//
#ifndef DOCXTOTXT_PARAGRAPHPARSER_H
#define DOCXTOTXT_PARAGRAPHPARSER_H

#include "tinyxml2.h"
#include "tinyxml.h"
#include <zip.h>
#include "iostream"
#include <map>
#include <vector>
#include <fstream>
#include "../headers/DrawingParser.h"

typedef struct {
    uint32_t flags;
    zip_t *input;
    std::ostream *output;
    std::string pathToDraws;
} options_t;

typedef struct {
    size_t docWidth;
    size_t docHeight;
    std::map<std::string, std::string> imageRelationship;
    std::map<std::string, std::string> hyperlinkRelationship;
    struct styles;
} docInfo_t;

typedef struct {//TODO pairs styleId/another struct or xml element
    std::map<std::string, std::string> characterStyles;
    std::map<std::string, std::string> numberingStyles;
    std::map<std::string, std::string> paragraphStyles;
    std::map<std::string, std::string> tableStyles;
} styles;
struct line {
    std::string text;
    size_t length = 0;
};
namespace paragraph {
    using namespace std;
    using namespace tinyxml2;

    enum language {
        enUS,
        ruRU
    };
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
    static map<string, int> languages = {{"en-US", enUS},
                                         {"ru-RU", ruRU}};

    class ParagraphParser {
    private:
        vector<line> paragraphBuffer;//TODO try wstring
        paragraphJustify justify;
        Drawing::DrawingParser drawingParser;
        options_t &options;
        docInfo_t &docInfo;

    public:
        ParagraphParser(docInfo_t &docInfo, options_t &options);

        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void parseHyperlink(XMLElement *properties);

        void addText(const string &text, language language);

        void setJustify(const string &justify);

        void setIndentation(XMLElement *element);

        void insertImage(size_t &height, size_t &width, const string &imageName = "");

        void writeResult();

        vector<line> getResult();

        void flush();
    };

}

#endif //DOCXTOTXT_PARAGRAPHPARSER_H
