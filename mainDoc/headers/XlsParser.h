//
// Created by borzor on 11/27/22.
//

#ifndef DOCXTOTXT_XLSPARSER_H
#define DOCXTOTXT_XLSPARSER_H

#include "ParagraphParser.h"

namespace xls {
    using namespace std;
    using namespace tinyxml2;

    class XlsParser {
    private:
        options_t &options;
        xlsInfo_t &xlsInfo;
        wstringConvert convertor;
    public:
        XlsParser(xlsInfo_t &xlsInfo, options_t &options);

        void parseSheets();

        static size_t getColumnWidth(const std::vector<columnSettings>& settings,size_t index);
    };

}
#endif //DOCXTOTXT_XLSPARSER_H
