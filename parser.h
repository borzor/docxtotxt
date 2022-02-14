//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_PARSER_H
#define DOCXTOTXT_PARSER_H
#define MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

#include "tinyxml2.h"
#include "tinyxml.h"
#include <zip.h>
#include "iostream"
#include <map>

namespace prsr {
    using namespace std;
    using namespace tinyxml2;

    class parser {
    private:
        int counter = 0;
        map<string, int> textProperties;
        map<string, int> paragraphProperties;
        string name;
        XMLDocument mainDoc;
        bool isInitialized;
        map<string, string> content_types;
    public:
        explicit parser(string name);

        void doInit(const string &path);

        void checkInit() const;

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseMainDoc();

        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void parseTable(XMLElement *table);

        void parseSection(XMLElement *section);
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
}

#endif //DOCXTOTXT_PARSER_H
