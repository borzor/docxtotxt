//
// Created by borzor on 12/12/22.
//

#include "../headers/DocParser.h"

namespace docxtotxt {
    DocParser::DocParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer) : docInfo(docInfo),
                                                                                         Parser(options, writer) {

    }

    void DocParser::parseFile() {
        writer.newLine();
        auto body = docInfo.body;
        for (auto &elem: body) {
            try {
                switch (elem.type) {
                    case par: {
                        writeParagraph(elem);
                        break;
                    }
                    case table:
                        writeTable(elem);
                        break;
                    case image:
                        writeImage(elem);
                        break;
                }
            } catch (exception &ignore) {
                writer.insertData(L"Произошла неожиданная ошибка, элемент проигнорирован", true, true);
            }
        }
        if ((options.flags >> 3) & 1) {
            for (auto &note: docInfo.notes) {
                writeNote(note);
            }
        }
    }

    void DocParser::writeParagraph(paragraph paragraph) {
        auto justify = paragraph.settings.justify;
        auto left = paragraph.settings.ind.left;
        auto right = paragraph.settings.ind.right;
        size_t firstLineLeft;
        if (paragraph.settings.ind.hanging == 0) {
            firstLineLeft = left + paragraph.settings.ind.firstLine;
        } else {
            firstLineLeft = left == 0 ? 0 : left - paragraph.settings.ind.hanging;
        }
        bool isFirstLine = true;
        auto before = paragraph.settings.spacing.before;
        auto after = paragraph.settings.spacing.after;
        for (auto &elem: paragraph.body) {
            auto currentSize = elem.length();
            if (!elem.empty()) {
                for (int i = 0; i < before; i++) {
                    writer.newLine();
                }
                while (currentSize != 0) {
                    auto availableBufferInLine = docInfo.docWidth - writer.getCurrentLength();
                    if (justify == left) {
                        if (isFirstLine) {
                            availableBufferInLine = availableBufferInLine - firstLineLeft - right;
                        } else {
                            availableBufferInLine = availableBufferInLine - left - right;
                        }
                    }
                    size_t ind = 0;
                    if (currentSize > availableBufferInLine) { // 167
                        auto indexLastElement = elem.find_last_of(L' ', availableBufferInLine);
                        if (indexLastElement == string::npos) {
//                            writer.newLine();
//                            availableBufferInLine = docInfo.docWidth - writer.getCurrentLength();
//                            indexLastElement = elem.find_last_of(L' ', availableBufferInLine);
                        }
                        switch (justify) {
                            case justify_t::left: {
                                ind = isFirstLine ? firstLineLeft : left;
                                break;
                            }
                            case justify_t::right: {
                                ind = docInfo.docWidth - indexLastElement;
                                break;
                            }
                            case justify_t::center: {
                                ind = (docInfo.docWidth - indexLastElement) / 2;
                                break;
                            }
                        }
                        writer.insertData(std::wstring(ind, L' ').append(elem.substr(0, indexLastElement)));
                        elem = elem.substr(indexLastElement + 1);
                        isFirstLine = false;
                        currentSize -= indexLastElement;
                    } else {
                        switch (justify) {
                            case justify_t::left: {
                                ind = isFirstLine ? firstLineLeft : left;
                                break;
                            }
                            case justify_t::right: {
                                ind = docInfo.docWidth - currentSize;
                                break;
                            }
                            case justify_t::center: {
                                ind = (docInfo.docWidth - currentSize) / 2;
                                break;
                            }
                        }
                        writer.insertData(std::wstring(ind, L' ').append(elem), false, false);
                        currentSize = 0;
                    }
                }
                for (int i = 0; i < after; i++) {
                    writer.newLine();
                }
            }
        }
        writer.newLine();
    }

    void DocParser::writeImage(const paragraph &paragraph) {
        for (auto &line: paragraph.body) {
            auto leftBorder = docInfo.docWidth > line.size() ? (docInfo.docWidth - line.size()) / 2 : 0;
            writer.insertData(std::wstring(leftBorder, L' ') + line);
        }
    }

    void DocParser::writeTable(const docxtotxt::paragraph &paragraph) {
        for (auto &line: paragraph.body) {
            auto leftBorder = docInfo.docWidth > line.size() ? (docInfo.docWidth - line.size()) / 2 : 0;
            writer.insertData(std::wstring(leftBorder, L' ') + line);
        }
    }

    void DocParser::writeNote(const note &note) {
        writer.insertData(L"Document notes:");
        if (note.type == noteType::footnote) {
            writer.insertData(std::wstring(L"{footnote" + to_wstring(note.id) + L"} - ") + note.text);
        } else {
            writer.insertData(std::wstring(L"{endnote" + to_wstring(note.id) + L"} - ") + note.text);
        }
    }


}
