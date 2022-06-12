//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H

#include "SectionParser.h"


namespace prsr {
    using namespace std;
    using namespace tinyxml2;

    class MainDocParser {
    private:
        XMLDocument mainDoc;
        bool isInitialized;
        map<string, string> content_types;
        options_t &options;
        docInfo_t &docInfo;
        wstringConvert convertor;

    public:

        explicit MainDocParser(options_t &options, ::docInfo_t &docInfo);

        void doInit();

        void checkInit() const;

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseMainDoc();

        void parseRelationships(XMLDocument *doc);

        void parseStyles(XMLDocument *doc);

        void addStyle(XMLElement *element);

        void setDefaultSettings(XMLElement *element);

        void insertHyperlinks();

        void saveImages();

    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
