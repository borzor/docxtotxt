//
// Created by boris on 2/23/22.
//

#include "../headers/TableParser.h"

namespace table {
    void TableParser::parseTable(XMLElement *table) {
        XMLElement *element = table->FirstChildElement();
        while (element != nullptr) {
            if (!strcmp(element->Value(), "w:tblPr")) {// table Properties
                parseTableProperties(element);
            } else if (!strcmp(element->Value(), "w:tblGrid")) {//TODO table width for each column IN twentieths of a point

            } else if (!strcmp(element->Value(), "w:tr")) {//table row
                parseTableRow(element);
            }
            element = element->NextSiblingElement();
        }
    }
    void TableParser::parseTableRow(XMLElement *row) {
        XMLElement *rowElement = row->FirstChildElement();
        while (rowElement != nullptr) {
            if (!strcmp(rowElement->Value(), "w:trPr")) {// Specifies the row-level properties for the row
                //TODO the most important attr - trHeight - the height of the row
            } else if (!strcmp(rowElement->Value(), "w:tblPrEx")) {//Specifies table properties for the row in place of the table properties specified in tblPr.
                //can scip, exceptions for table-level properties
            } else if (!strcmp(rowElement->Value(), "w:tc")) {//Specifies a table cell
                XMLElement *tcElement = rowElement->FirstChildElement();
                while (tcElement != nullptr){
                    if (!strcmp(tcElement->Value(), "w:p")) {//
                        //TODO paragraph parser...
                    } else if (!strcmp(tcElement->Value(), "w:tbl")) {//
                        // for the table in table...
                    } else if (!strcmp(tcElement->Value(), "w:tcPr")) {//Specifies a table cell
                        // a lot of boring attrs, looks like skip
                    }
                    tcElement = tcElement->NextSiblingElement();
                }
            }
            rowElement = rowElement->NextSiblingElement();
        }

    }
    void TableParser::parseTableProperties(XMLElement *element) {
        XMLElement *tableProperty = element->FirstChildElement();
        while (tableProperty != nullptr) {
            switch (tableProperties[tableProperty->Value()]) {
                case jc:
                    break;
                case shd:
                    break;
                case tblBorders:
                    break;
                case tblCaption:
                    break;
                case tblCellMar:
                    break;
                case tblCellSpacing:
                    break;
                case tblInd:
                    break;
                case tblLayout:
                    break;
                case tblLook:
                    break;
                case tblOverlap:
                    break;
                case tblpPr:
                    break;
                case tblStyle:
                    break;
                case tblStyleColBandSize:
                    break;
                case tblStyleRowBandSize:
                    break;
                case tblW:
                    //TODO table WIDTH
                    break;

            }
            tableProperty = tableProperty->NextSiblingElement();
        }
        free(tableProperty);

    }




}

