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
#include "../headers/DrawingParser.h"

typedef struct {
    uint32_t flags;
    zip_t *input;
    std::ofstream output;
    std::string pathToDraws;
} options_t;

namespace paragraph {
    using namespace std;
    using namespace tinyxml2;
    struct line {
        string text;
        size_t length = 0;
    };
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
        size_t amountOfCharacters;
        Drawing::DrawingParser drawingParser;
        map<string, string> &imageRelationship;
        options_t options;


    public:
        ParagraphParser(size_t amountOfCharacters, options_t options,
                        map<string, string> &imageRelationship);

        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void parseHyperlink(XMLElement *properties);

        void addText(const string &text, language language);

        void setJustify(const string &justify);

        void setIndentation(XMLElement *element);

        void insertImage(size_t &height, size_t &width, const string &imageName = "");

        void writeResult();

        void flush();
    };

}

