//
// Created by boris on 2/20/22.
//
#ifndef DOCXTOTXT_PARAGRAPHPARSER_H
#define DOCXTOTXT_PARAGRAPHPARSER_H

#include "tinyxml2.h"
#include <zip.h>
#include "iostream"
#include <map>
#include <vector>
#include <fstream>
#include <locale>
#include <codecvt>
#include "../headers/DrawingParser.h"
#include "../ParserCommons/DocumentCommons.h"

namespace paragraph {
    using namespace std;
    using namespace tinyxml2;

    class ParagraphParser {
    private:
        options_t &options;
        docInfo_t &docInfo;
        wstring paragraphBuffer;
        paragraphSettings settings;
        Drawing::DrawingParser drawingParser;
        wstringConvert convertor;
    public:
        ParagraphParser(docInfo_t &docInfo, options_t &options);

        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void parseHyperlink(XMLElement *properties);

        void insertImage(size_t &height, size_t &width, const string &imageName = "");

        void writeResult();

        wstring getResult();

        void flush();

        static void setJustify(XMLElement *element, paragraphJustify &settings);

        static void setIndentation(XMLElement *element, indentation &settings);

        static void setSpacing(XMLElement *element, lineSpacing &settings);

        static void setTabulation(XMLElement *element, std::queue<tabulation> &settings);

    };

}

#endif //DOCXTOTXT_PARAGRAPHPARSER_H
