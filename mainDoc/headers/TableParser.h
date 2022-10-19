//
// Created by boris on 2/23/22.
//

#ifndef DOCXTOTXT_TABLEPARSER_H
#define DOCXTOTXT_TABLEPARSER_H

#include <iostream>
#include "tinyxml2.h"
#include "../ParserCommons/DocumentCommons.h"
#include "ParagraphParser.h"
#include <map>

static std::map<std::string, int> tablePropertiesMap =
        {{"w:jc",                  tblJc},
         {"w:shd",                 tblJshd},
         {"w:tblBorders",          tblBorders},
         {"w:tblCaption",          tblCaption},
         {"w:tblCellMar",          tblCellMar},
         {"w:tblCellSpacing",      tblCellSpacing},
         {"w:tblInd",              tblInd},
         {"w:tblLayout",           tblLayout},
         {"w:tblLook",             tblLook},
         {"w:tblOverlap",          tblOverlap},
         {"w:tblpPr",              tblpPr},
         {"w:tblStyle",            tblStyle},
         {"w:tblStyleColBandSize", tblStyleColBandSize},
         {"w:tblStyleRowBandSize", tblStyleRowBandSize},
         {"w:tblW",                tblW}};

typedef struct {
    size_t width;
    size_t gridSpan;
    std::wstring text;
} cell;

typedef struct {
    size_t columnAmount;
    tableProperties settings;
    std::vector<std::size_t> tblGrids;
    std::vector<std::vector<cell>> grid;
} tableSetting;

namespace table {
    using namespace std;
    using namespace tinyxml2;

    class TableParser {
    private:
        size_t currentColumn = 0;
        size_t line = 0;
        options_t &options;
        docInfo_t &docInfo;
        tableSetting table;
    public:
        TableParser(docInfo_t &docInfo, options_t &options);

        void parseTable(XMLElement *table);

        void parseTableProperties(XMLElement *properties);

        void parseTableRow(XMLElement *row);

        void parseTableRowProperties(XMLElement *row);

        void parseTableGrid(XMLElement *grid);

        void getTableCellProperties(XMLElement *cell);

        void getTableCellParagraph(XMLElement *cell, paragraph::ParagraphParser &paragraphParser);

        void insertTable();

        void flush();

        static void setJustify(XMLElement *element, tableJustify &settings);

        static void setIndentation(XMLElement *element, tableIndent &settings);

        static void setFloatingSettings(XMLElement *element, floatingSettings &settings);
    };
}

#endif //DOCXTOTXT_TABLEPARSER_H
