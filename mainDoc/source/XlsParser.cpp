//
// Created by borzor on 11/27/22.
//

#include "../headers/XlsParser.h"

namespace docxtotxt {
    XlsParser::XlsParser(xlsInfo_t &xlsInfo, options_t &options, BufferWriter &writer) : xlsInfo(xlsInfo),
                                                                                         Parser(options, writer) {

    }

    void XlsParser::parseFile() {
        for (auto &sh: xlsInfo.worksheets) {
            insertSheetMetadata(sh);
            auto array = sh.sheetArray;
            if (array.empty())
                continue;
            auto col = sh.col;
            size_t numberOfColumn = 0;
            for (auto &i: array)
                numberOfColumn = std::max(numberOfColumn, i.back().cellNumber + 1);
            std::vector<size_t> columnSize(numberOfColumn, 0);
            size_t tableWidth = numberOfColumn - 1;
            for (int column = 0; column < columnSize.size(); column++) {
                auto width = getColumnWidth(col, column);
                if (width < 10)
                    width = 10;
                columnSize[column] = width;
                tableWidth += width;
            }
            writer.insertData(std::wstring(L"+").append(tableWidth, L'-').append(L"+"), false, false);
            size_t line = 0;
            while (line < array.size()) {
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
                writer.insertData(L"|", true, false);
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
                            writer.insertData(resultText, false, false);
                            writer.insertData(
                                    std::wstring(indexLastElement == string::npos ? 0 : charInCell - indexLastElement,
                                                 L' '), false, false);
                        } else {
                            resultText = text;
                            writer.insertData(resultText, false, false);
                            writer.insertData(std::wstring(charInCell - text.size(), L' '), false, false);
                        }
                        array[line][currentIndex].text.erase(0, resultText.size() + 1);
                        writer.insertData(L"|", false, false);
                    } else {
                        writer.insertData(std::wstring(charInCell == 1 ? 0 : charInCell, L' ') + L"|", false, false);
                    }
                    currentIndex++;
                }
                if (lineDone) {
                    line++;
                    writer.insertData(std::wstring(L"+").append(tableWidth, L'-').append(L"+"), true, false);
                }
            }
            writer.newLine();
        }
    }

    size_t XlsParser::getColumnWidth(const std::vector<columnSettings> &settings, size_t index) {
        for (auto &setting: settings) {
            if (setting.startInd <= index && setting.endIndInd >= index) {
                return setting.width;
            }
        }
        return 0;
    }

    void XlsParser::insertSheetMetadata(const sheet &sheet) {
        writer.insertData(L"Sheet info: ", true);
        writer.insertData(
                L"SheetName - " + sheet.sheetName + (sheet.state.empty() ? L"" : L" , State - " + sheet.state));
        for (const auto &kv: sheet.relations.drawing) {
            auto draw = std::find(xlsInfo.draws.begin(), xlsInfo.draws.end(), kv.second);
            writer.insertData(L"Sheet image info: ", false, false);
            for (const auto &qw: draw->relations.imageRelationship) {
                writer.insertData(writer.convertString(qw.second), false, false);
                if ((options.flags >> 1) & 1) {
                    string image = ", saved to path: " + options.pathToDraws + '/' + qw.second;
                    writer.insertData(writer.convertString(image));
                }
            }
        }
    }

}
