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
#include "../ParserCommons/DocumentCommons.h"
#include "../ParserCommons/FileSpecificCommons/DocCommons.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;

    class ParagraphParser {
    private:
        options_t &options;
        docInfo_t &docInfo;
        BufferWriter &writer;
        wstring paragraphBuffer;
        paragraphSettings settings;
        wstringConvert convertor;
    public:
        ParagraphParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer);

        void parseParagraph(XMLElement *paragraph);

        void parseParagraphProperties(XMLElement *properties);

        void parseTextProperties(XMLElement *properties);

        void parseHyperlink(XMLElement *properties);

        void insertTabulation(paragraphSettings tab);

        void insertImage(picture pic);

        static void setJustify(XMLElement *element, paragraphJustify &settings);

        static void setIndentation(XMLElement *element, indentation &settings);

        static void setSpacing(XMLElement *element, lineSpacing &settings);

        static void setTabulation(XMLElement *element, std::queue<tabulation> &settings);

        void writeResult();

        wstring getResult();

        void flush();
    };

}

#endif //DOCXTOTXT_PARAGRAPHPARSER_H
