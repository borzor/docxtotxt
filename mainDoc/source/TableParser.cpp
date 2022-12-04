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
            } else if (!strcmp(element->Value(), "w:tblGrid")) {
                parseTableGrid(element);
            } else if (!strcmp(element->Value(), "w:tr")) {//table row
                parseTableRow(element);
                currentColumn = 0;
                line += 1;
            }
            element = element->NextSiblingElement();
        }
    }

    void TableParser::parseTableRow(XMLElement *row) {
        XMLElement *rowElement = row->FirstChildElement();
        paragraph::ParagraphParser paragraphParser(docInfo, options);
        table.grid.emplace_back();
        while (rowElement != nullptr) {
            if (!strcmp(rowElement->Value(), "w:trPr")) {
                parseTableRowProperties(rowElement);
            } else if (!strcmp(rowElement->Value(), "w:tblPrEx")) {
                //Specifies table properties for the row in place of the table properties specified in tblPr.
                //can scip, exceptions for table-level properties
            } else if (!strcmp(rowElement->Value(), "w:tc")) {//Specifies a table cell
                XMLElement *tcElement = rowElement->FirstChildElement();
                while (tcElement != nullptr) {
                    if (!strcmp(tcElement->Value(), "w:p")) {//
                        getTableCellParagraph(tcElement, paragraphParser);
                    } else if (!strcmp(tcElement->Value(), "w:tbl")) {//
                        // for the table in table...
                    } else if (!strcmp(tcElement->Value(), "w:tcPr")) {//Specifies a table cell
                        getTableCellProperties(tcElement);
                    }
                    tcElement = tcElement->NextSiblingElement();
                }
                currentColumn += table.grid[line].back().gridSpan;

            }
            rowElement = rowElement->NextSiblingElement();
        }
    }

    void TableParser::parseTableProperties(XMLElement *properties) {
        auto tblStyle = properties->FirstChildElement("w:tblStyle");
        if (tblStyle == nullptr) {
            table.settings = docInfo.styles.defaultStyles.table;
        } else {
            table.settings = docInfo.styles.tableStyles[tblStyle->Attribute("w:val")];
        }
        XMLElement *tableProperty = properties->FirstChildElement();
        while (tableProperty != nullptr) {
            switch (tablePropertiesMap[tableProperty->Value()]) {
                case tblJc:
                    if (tableProperty->FirstAttribute() != nullptr)
                        //setJustify(tableProperty, table.settings.justify);
                        break;
                case tblJshd:
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
                    if (tableProperty->FirstAttribute() != nullptr)
                        setIndentation(tableProperty, table.settings.ind);
                    break;
                case tblLayout:
                    break;
                case tblLook:
                    break;
                case tblOverlap://TODO check visibility
                    break;
                case tblpPr:
                    if (tableProperty->FirstAttribute() != nullptr)
                        setFloatingSettings(tableProperty, table.settings.floatTable);
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
            table.tblGrids.emplace_back(std::stoi(gridCol->Attribute("w:w")));
            gridCol = gridCol->NextSiblingElement();
        }
        table.columnAmount = table.tblGrids.size();
        free(gridCol);
    }

    TableParser::TableParser(docInfo_t &docInfo, options_t &options) : options(options), docInfo(docInfo) {
    }

    void TableParser::insertTable() {
        this->currentColumn = 0;
        this->line = 0;
        size_t tableWidth = 0;
        for (auto &cell: table.grid[line]) {
            tableWidth += cell.width;
        }
        size_t minColumn = table.columnAmount;
        for (auto & i : table.grid) {
            minColumn = min(minColumn, i.size());
        }
        auto tableSize = tableWidth;
        tableSize += (minColumn + 1); //column size + amount of columns
        auto mediumLine = tableSize - 2;
        auto leftBorder = (docInfo.docWidth - tableSize) / 2;
        docInfo.resultBuffer.buffer.back().append(
                std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—')).append(1, L'+');
        while (line < table.grid.size()) {
            bool lineDone = true;
            currentColumn = 0;
            addLine(docInfo.resultBuffer);
            docInfo.resultBuffer.buffer.back().append(std::wstring(leftBorder, L' ') + L"|");
            while (currentColumn < table.grid[line].size()) {
                auto charInCell = table.grid[line][currentColumn].width;
                if (!table.grid[line][currentColumn].text.empty()) {
                    auto text = table.grid[line][currentColumn].text;
                    std::wstring resultText;
                    auto indexLastElement = text.find_last_of(L' ', charInCell);
                    if (charInCell < text.size()) {
                        lineDone = false;
                        if (indexLastElement == string::npos) {
                            resultText = text.substr(0, charInCell);
                        } else {
                            resultText = text.substr(0, indexLastElement);
                        }
                        docInfo.resultBuffer.buffer.back().append(resultText);
                        docInfo.resultBuffer.buffer.back().append(
                                indexLastElement == string::npos ? 0 : charInCell - indexLastElement, L' ');
                    } else {
                        resultText = text;
                        docInfo.resultBuffer.buffer.back().append(resultText);
                        docInfo.resultBuffer.buffer.back().append(charInCell - text.size(), L' ');
                    }
                    table.grid[line][currentColumn].text.erase(0, resultText.size() + 1);
                    docInfo.resultBuffer.buffer.back().append(1, L'|');
                } else {
                    docInfo.resultBuffer.buffer.back().append(charInCell == 1 ? 0 : charInCell, L' ').append(1, L'|');
                }
                currentColumn++;
            }
            if (lineDone) {
                line++;
                addLine(docInfo.resultBuffer);
                docInfo.resultBuffer.buffer.back().append(
                        std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—')).append(1, L'+');
            }
        }
        addLine(docInfo.resultBuffer);
    }

    void TableParser::flush() {
        memset(&table.settings, 0, sizeof(table.settings));
        table.grid.clear();
        table.tblGrids.clear();
        currentColumn = 0;
        line = 0;
    }

    void TableParser::setJustify(XMLElement *jc, tableJustify &settings) {
        if (jc != nullptr) {
            string justify = jc->Attribute("w:w");
            if (!strcmp(justify.c_str(), "start"))
                settings = tableJustify::start;
            else if (!strcmp(justify.c_str(), "end"))
                settings = tableJustify::end;
            else if (!strcmp(justify.c_str(), "center"))
                settings = tableJustify::center;
        }
    }

    void TableParser::setIndentation(XMLElement *ind, tableIndent &settings) {
        if (ind != nullptr) {
            auto tblInd = ind->Attribute("w:w");
            settings.ind = atoi(tblInd);
        }
    }

    void TableParser::setFloatingSettings(XMLElement *element, floatingSettings &settings) {

    }


    void TableParser::parseTableRowProperties(XMLElement *row) {//TODO trHeight - the height of the row

    }

    void TableParser::getTableCellProperties(XMLElement *element) {
        XMLElement *gridCol = element->FirstChildElement();
        size_t gridSpan = 1;
        cell tmpCell;
//        tmpCell.width = table.tblGrids[currentColumn];
        while (gridCol != nullptr) {
            if (!strcmp(gridCol->Value(), "w:tcW")) {
                auto size = gridCol->Attribute("w:w");
                if (size != nullptr) {
                    tmpCell.width = atoi(size) / TWIP_TO_CHARACTER;
                }
            } else if (!strcmp(gridCol->Value(), "w:gridSpan")) {
                gridSpan = atoi(gridCol->Attribute("w:val"));
            }
            gridCol = gridCol->NextSiblingElement();
        }
        tmpCell.gridSpan = gridSpan;
        table.grid[line].push_back(tmpCell);
        free(gridCol);
    }

    void TableParser::getTableCellParagraph(XMLElement *tcElement, paragraph::ParagraphParser &paragraphParser) {
        paragraphParser.parseParagraph(tcElement);
        auto text = paragraphParser.getResult();
        table.grid[line].back().text += text;
        paragraphParser.flush();
    }


}

