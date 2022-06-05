//
// Created by boris on 2/23/22.
//

#ifndef DOCXTOTXT_TABLEPARSER_H
#define DOCXTOTXT_TABLEPARSER_H

#include <iostream>
#include "tinyxml2.h"
#include "tinyxml.h"
#include "ParagraphParser.h"
#include <map>


namespace table {
    using namespace std;
    using namespace tinyxml2;
    enum tableProperty {
        jc,
        shd,
        tblBorders,
        tblCaption,
        tblCellMar,
        tblCellSpacing,
        tblInd,
        tblLayout,
        tblLook,
        tblOverlap,
        tblpPr,
        tblStyle,
        tblStyleColBandSize,
        tblStyleRowBandSize,
        tblW
    };
    static map<string, int> tableProperties = {{"w:jc",                  jc},
                                               {"w:shd",                 shd},
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

    class TableParser {
    private:
        options_t &options;
        docInfo_t &docInfo;
    public:
        TableParser(docInfo_t &docInfo, options_t &options);

        void parseTable(XMLElement *table);

        void parseTableProperties(XMLElement *properties);

        void parseTableRow(XMLElement *row);

        void parseTableGrid(XMLElement *grid);
    };
}

#endif //DOCXTOTXT_TABLEPARSER_H
