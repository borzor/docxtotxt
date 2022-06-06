//
// Created by borzor on 11/27/22.
//

#ifndef DOCXTOTXT_XLSPARSER_H
#define DOCXTOTXT_XLSPARSER_H

#include "../ParserCommons/FileSpecificCommons/XlsCommons.h"

namespace xls {
    using namespace std;
    using namespace tinyxml2;

    class XlsParser {
    private:
        options_t &options;
        xlsInfo_t &xlsInfo;
        wstringConvert convertor;

        static size_t getColumnWidth(const std::vector<columnSettings> &settings, size_t index);

        void insertSlideMetaData(const sheet& sheet);

//       void insertSheetMetaData(slideInsertInfo relations);

    public:
        XlsParser(xlsInfo_t &xlsInfo, options_t &options);

        void parseSheets();

    };

}
#endif //DOCXTOTXT_XLSPARSER_H
