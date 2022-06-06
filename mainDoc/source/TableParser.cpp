//
// Created by boris on 2/23/22.
//

#include "../headers/TableParser.h"
#include "../headers/SectionParser.h"
#include <iomanip>
namespace table {
    void TableParser::parseTable(XMLElement *table) {
        XMLElement *element = table->FirstChildElement();
        while (element != nullptr) {
            if (!strcmp(element->Value(), "w:tblPr")) {// table Properties
                parseTableProperties(element);
            } else if (!strcmp(element->Value(),
                               "w:tblGrid")) {//TODO table width for each column IN twentieths of a point
                parseTableGrid(element);
            } else if (!strcmp(element->Value(), "w:tr")) {//table row
                parseTableRow(element);
            }
            element = element->NextSiblingElement();
        }
        insertTable();
        flush();
    }

    void TableParser::parseTableRow(XMLElement *row) {
        XMLElement *rowElement = row->FirstChildElement();
        auto paragraphParser = paragraph::ParagraphParser(docInfo, options);
        while (rowElement != nullptr) {
            if (!strcmp(rowElement->Value(), "w:trPr")) {// Specifies the row-level properties for the row
                //TODO the most important attr - trHeight - the height of the row
            } else if (!strcmp(rowElement->Value(),
                               "w:tblPrEx")) {//Specifies table properties for the row in place of the table properties specified in tblPr.
                //can scip, exceptions for table-level properties
            } else if (!strcmp(rowElement->Value(), "w:tc")) {//Specifies a table cell
                XMLElement *tcElement = rowElement->FirstChildElement();
                while (tcElement != nullptr) {
                    if (!strcmp(tcElement->Value(), "w:p")) {//
                        paragraphParser.parseParagraph(tcElement);
                        tableData.push_back(paragraphParser.getResult().back());
                        paragraphParser.flush();
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

    void TableParser::parseTableGrid(XMLElement *tblGrid) {
        XMLElement *gridCol = tblGrid->FirstChildElement();
        while (gridCol != nullptr) {
            tblGrids.emplace_back(std::stoi(gridCol->Attribute("w:w")));
            gridCol = gridCol->NextSiblingElement();
        }
        free(gridCol);
    }

    TableParser::TableParser(docInfo_t &docInfo, options_t &options) : options(options), docInfo(docInfo) {

    }

    void TableParser::insertTable() {
        size_t width = 0;
        for (const auto &item: tblGrids) {
            width += item;
        }
        auto leftBorder = (docInfo.docWidth - (width / TWIP_TO_CHARACTER)) / 2;
        auto counter = 0;
        while (counter != tableData.size()) {
            line tmp;
            tmp.text.insert(0, leftBorder, ' ');
            auto currentOffset = leftBorder;
            for (const auto &item: tblGrids) {
                auto characterLength = tableData[counter].length;
                auto data = tableData[counter].text;
                auto charactersInCell = item / TWIP_TO_CHARACTER;
                while(characterLength != 0){
                    if(characterLength > charactersInCell){
                        tmp.text.append(tableData[counter].text);
                        tableData[counter].text;
                        resultTable.push_back(tmp);
                        tmp.text = "";
                        tmp.length = 0;
                        characterLength-=charactersInCell;
                    } else{

                    }
                }

                auto offset = charactersInCell - tableData[counter].length;
                tmp.text.append(offset / 2, ' ');
                tmp.text.append(tableData[counter].text);
                tmp.text.append(offset / 2, ' ');
                currentOffset += item;
                counter++;
            }
            tmp.length += docInfo.docWidth;
            resultTable.push_back(tmp);
        }
        for (auto &s: resultTable) {
            *options.output << setw((strlen(s.text.c_str()) + docInfo.docWidth / 2 - s.length / 2))
                            << s.text
                            << endl;
        }

    }

    void TableParser::flush() {
        tblGrids.clear();
        tableData.clear();
        resultTable.clear();
    }


}

