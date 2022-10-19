//
// Created by borzor on 10/18/22.
//

#ifndef DOCXTOTXT_DOCUMENTPROPERTIESPARSER_H
#define DOCXTOTXT_DOCUMENTPROPERTIESPARSER_H

#include "../ParserCommons/DocumentCommons.h"
#include "../ParserCommons/DocumentMetaDataCommons.h"

namespace document {
    class DocumentPropertiesParser {
    private:
        static void parseWordApp(tinyxml2::XMLDocument &doc, wordMetaData &data);
        static void parseWordCore(tinyxml2::XMLDocument &doc, wordMetaData &data);
        static void printWordData(const options_t& options, wordMetaData &data);

        static void parsePptApp(tinyxml2::XMLDocument &doc, pptMetaData &data);
        static void parsePptCore(tinyxml2::XMLDocument &doc,pptMetaData &data);
        static void printPptData(const options_t& options, pptMetaData &data);

        static void parseXlsxApp(tinyxml2::XMLDocument &doc, xlsxMetaData &data);
        static void parseXlsxCore(tinyxml2::XMLDocument &doc,xlsxMetaData &data);
        static void printXlsxData(const options_t& options, xlsxMetaData &data);
    public:
        static void printData(const options_t& options);

    };
}

#endif //DOCXTOTXT_DOCUMENTPROPERTIESPARSER_H
