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
        string name;
        XMLDocument mainDoc;
        bool isInitialized;
        map<string, string> content_types;
    public:
        explicit parser(const std::string &name);
        void doInit(const std::string &path);
        void checkInit() const;
        void parseContentTypes(XMLDocument *xmlDocument);
        void parseMainDoc();
        void parseParagraph(XMLElement *paragraph);
        void parseTable(XMLElement *table);
        void parseSection(XMLElement *section);
    };
}

#endif //DOCXTOTXT_PARSER_H
