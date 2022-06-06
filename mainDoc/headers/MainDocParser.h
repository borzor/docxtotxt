//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H


#include "../ParserCommons/DocumentCommons.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;

    class MainDocParser {
    private:
        options_t &options;
        wstringConvert convertor;
    public:
        explicit MainDocParser(options_t &options);

        void parseFile();

        void printResult(buffer &resultBuffer) const;

        void insertHyperlinks(std::map<std::string, std::string> hyperlinkRelationship);

        void saveImages(const std::map<std::string, std::string> &imageRelationship) const;
    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
