//
// Created by borzor on 11/27/22.
//

#ifndef DOCXTOTXT_XLSPARSER_H
#define DOCXTOTXT_XLSPARSER_H

#include "../ParserCommons/FileSpecificCommons/XlsCommons.h"
#include "../ParserCommons/BufferWriter.h"
#include "Parser.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;
    /*!
	\brief Реализация парсера, обрабатывающая .xlsx формат
    */
    class XlsParser : public Parser{
    private:
        ///Настройки специфичные для файла .pptx формата
        xlsInfo_t &xlsInfo;

        static size_t getColumnWidth(const std::vector<columnSettings> &settings, size_t index);

        void insertSheetMetadata(const sheet &sheet);

    public:
        XlsParser(xlsInfo_t &xlsInfo, options_t &options, BufferWriter &writer);

        void parseFile() override;

    };

}
#endif //DOCXTOTXT_XLSPARSER_H
