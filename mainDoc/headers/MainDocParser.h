//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H
#define MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#include "ParagraphParser.h"
#include "SectionParser.h"


namespace prsr {
    using namespace std;
    using namespace tinyxml2;

    class MainDocParser {
    private:
        XMLDocument mainDoc;
        XMLDocument imagesDoc;
        map<string, string> imageRelationship;
        bool isInitialized;
        map<string, string> content_types;
        options_t &options;
    public:

        explicit MainDocParser(options_t &options);

        void doInit();

        void checkInit() const;

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseMainDoc();

        void parseImageDoc(XMLDocument *doc);

    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
