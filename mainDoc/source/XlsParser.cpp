//
// Created by borzor on 11/27/22.
//

#include "../headers/XlsParser.h"

namespace xls {
    XlsParser::XlsParser(xlsInfo_t &xlsInfo, options_t &options) : xlsInfo(xlsInfo), options(options) {

    }

    void XlsParser::parseSheets() {
        for (auto &sh: xlsInfo.worksheets) {
            addLine(xlsInfo.resultBuffer);
            auto sheetName = sh.second.sheetName;
            auto state = sh.second.state;
            auto array = sh.second.sheetArray;
            auto col = sh.second.col;
            size_t numberOfColumn = 0;
            xlsInfo.resultBuffer.buffer.back().append(L"Sheet info: ");
            addLine(xlsInfo.resultBuffer);
            xlsInfo.resultBuffer.buffer.back().append(L"SheetName - ").append(sheetName);
            if(!state.empty())
                xlsInfo.resultBuffer.buffer.back().append(L", State - ").append(state);
            addLine(xlsInfo.resultBuffer);
            for (auto &i: array)
                numberOfColumn = std::max(numberOfColumn, i.back().cellNumber + 1);
            std::vector<size_t> columnSize(numberOfColumn, 0);
            size_t tableWidth = numberOfColumn - 1;
            for (int column = 0; column < columnSize.size(); column++) {
                auto width = getColumnWidth(col, column);
                if(width < 10)
                    width = 10;
                columnSize[column] = width;
                tableWidth += width;
            }
            xlsInfo.resultBuffer.buffer.back().append(
                    std::wstring(1, L'+').append(tableWidth, L'—')).append(1, L'+');
            size_t line = 0;
            while (line < array.size()) {
                addLine(xlsInfo.resultBuffer);
                bool lineDone = true;
                size_t currentIndex = 0;
                for (int column = 0; column < numberOfColumn; column++) {
                    if (array[line][currentIndex].cellNumber == currentIndex) {
                        currentIndex++;
                    } else {
                        sheetCell tmpCell;
                        tmpCell.cellNumber = currentIndex;
                        tmpCell.text = L"";
                        array[line].insert(array[line].begin() + currentIndex, tmpCell);
                    }
                    if (column > array[line].size() - 1) {
                        sheetCell tmpCell;
                        tmpCell.text = L"";
                        tmpCell.cellNumber = column;
                        array[line].emplace_back(tmpCell);
                    }
                }
                currentIndex = 0;
                xlsInfo.resultBuffer.buffer.back().append(L"|");
                while (currentIndex < numberOfColumn) {
                    auto charInCell = columnSize[currentIndex];
                    auto cell = array[line][currentIndex];
                    if (!cell.text.empty()) {
                        auto text = cell.text;
                        std::wstring resultText;
                        auto indexLastElement = text.find_last_of(L' ', charInCell);
                        if (charInCell < text.size()) {
                            lineDone = false;
                            if (indexLastElement == string::npos) {
                                resultText = text.substr(0, charInCell);
                            } else {
                                resultText = text.substr(0, indexLastElement);
                            }
                            xlsInfo.resultBuffer.buffer.back().append(resultText);
                            xlsInfo.resultBuffer.buffer.back().append(
                                    indexLastElement == string::npos ? 0 : charInCell - indexLastElement, L' ');
                        } else {

                            resultText = text;
                            xlsInfo.resultBuffer.buffer.back().append(resultText);
                            xlsInfo.resultBuffer.buffer.back().append(charInCell - text.size(), L' ');
                        }
                        array[line][currentIndex].text.erase(0, resultText.size() + 1);
                        xlsInfo.resultBuffer.buffer.back().append(1, L'|');
                    } else {
                        xlsInfo.resultBuffer.buffer.back().append(charInCell == 1 ? 0 : charInCell, L' ').append(1,
                                                                                                                 L'|');
                    }
                    currentIndex++;
                }
                if (lineDone) {
                    line++;
                    addLine(xlsInfo.resultBuffer);
                    xlsInfo.resultBuffer.buffer.back().append(
                            std::wstring(1, L'+').append(tableWidth, L'—')).append(1, L'+');;
                }
            }
            addLine(xlsInfo.resultBuffer);
        }
    }

    size_t XlsParser::getColumnWidth(const std::vector<columnSettings>& settings, size_t index) {
        for (auto &setting: settings) {
            if (setting.startInd <= index && setting.endIndInd >= index) {
                return setting.width;
            }
        }
        return 0;
    }

}
