//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H
#define MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#include "SectionParser.h"


namespace prsr {
    using namespace std;
    using namespace tinyxml2;

    class MainDocParser {
    private:
        XMLDocument mainDoc;
        XMLDocument relationsDoc;
        bool isInitialized;
        map<string, string> content_types;
        options_t &options;
        docInfo_t &docInfo;

    public:

        explicit MainDocParser(options_t &options, ::docInfo_t &docInfo);

        void doInit();

        void checkInit() const;

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseMainDoc();

        void parseRelationships(XMLDocument *doc);

        void insertHyperlinks();

    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
