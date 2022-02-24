//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H
#define MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

#include "ParagraphParser.h"
#include "SectionParser.h"

namespace prsr {
    using namespace std;
    using namespace tinyxml2;

    class MainDocParser {
    private:
        int counter = 0;
        map<string, int> textProperties;
        map<string, int> paragraphProperties;
        string name;
        XMLDocument mainDoc;
        bool isInitialized;
        map<string, string> content_types;
        bool toFile;
    public:
        explicit MainDocParser(string name);

        void doInit(const string &path);

        void checkInit() const;

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseMainDoc();

        void parseTable(XMLElement *table);

    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
