//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_DOCPARSER_H
#define DOCXTOTXT_DOCPARSER_H

#include "ParagraphParser.h"
#include "Parser.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;
    /*!
	\brief Реализация парсера, обрабатывающая .docx формат
    */
    class DocParser : public Parser {
    private:
        ///Настройки специфичные для файла .docx формата
        docInfo_t &docInfo;

        void parseSection(XMLElement *section);

    public:
        DocParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer);

        void parseFile() override{};

        void parseFile(XMLDocument *mainFile);
    };

}
#endif //DOCXTOTXT_DOCPARSER_H
