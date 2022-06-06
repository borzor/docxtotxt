//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_DOCPARSER_H
#define DOCXTOTXT_DOCPARSER_H

#include "ParagraphParser.h"

namespace doc {
    using namespace std;
    using namespace tinyxml2;

    class DocParser {
    private:
        options_t &options;
        docInfo_t &docInfo;

        void parseSection(XMLElement *section);

    public:
        DocParser(docInfo_t &docInfo, options_t &options);

        void parseMainFile(XMLDocument *mainFile);
    };

}
#endif //DOCXTOTXT_DOCPARSER_H
