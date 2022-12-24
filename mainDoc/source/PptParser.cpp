//
// Created by borzor on 12/12/22.
//

#include "../headers/PptParser.h"
#include "algorithm"

namespace docxtotxt {
    PptParser::PptParser(pptInfo_t &pptInfo, options_t &options, BufferWriter &writer) : pptInfo(pptInfo),
                                                                                         Parser(options, writer) {
    }

    void PptParser::parseFile() {
        prepareSlides();
        writer.newLine();
        for (auto &obj: slideInsertData) {
            insertSlideMetadata(obj);
            insertSlide(obj.insertObjects);
        }
    }

    void PptParser::prepareSlides() {
        size_t slideNum = 1;
        for (const auto &slide: pptInfo.slides) {
            std::vector<insertObject> slideObjects;
            std::vector<insertImage> insertImages;
            for (const auto &presObj: slide.objects) {
                if (!presObj.paragraph.empty()) {
                    insertObject obj;
                    obj.offset = presObj.objectInfo.offsetX / pptInfo.settings.widthCoefficient;
                    obj.startLine = presObj.objectInfo.offsetY / pptInfo.settings.heightCoefficient;
                    obj.paragraph = presObj.paragraph;
                    obj.length = presObj.objectInfo.objectSizeX / pptInfo.settings.widthCoefficient;
                    obj.inProgress = false;
                    slideObjects.emplace_back(obj);
                }
            }
            for (const auto &table: slide.tables) {
                size_t counter = 0;
                size_t lineLength = 0;
                std::for_each(table.gridColSize.begin(), table.gridColSize.end(), [&](int n) {
                    lineLength += n;
                });
                for (auto &line: table.table) {
                    bool repeatLine;
                    auto tmpLine = line;
                    do {
                        insertObject obj;
                        repeatLine = false;
                        counter++;
                        textBody text;
                        for (auto &cell: tmpLine) {
                            auto colNum = 0;
                            for (auto &txt: cell) {
                                if (txt.text.length() > table.gridColSize[colNum]) {
//                                    text.text.append(txt.text.substr(0, table.gridColSize[colNum] - 1));
//                                    txt.text = txt.text.substr(table.gridColSize[colNum] - 1);
//                                    text.text.append(1, ' ');
//                                    repeatLine = true;
                                } else {
                                    text.text.append(txt.text);
                                    text.text.append(table.gridColSize[colNum] - txt.text.length(), ' ');
                                    txt.text = L"";
                                }
                                colNum++;
                            }
                        }
                        text.align = l;
                        obj.length = lineLength;
                        obj.offset = table.objectInfo.offsetX / pptInfo.settings.widthCoefficient;
                        obj.startLine = table.objectInfo.offsetY / pptInfo.settings.heightCoefficient + counter;
                        obj.paragraph.emplace_back(text);
                        obj.inProgress = false;
                        slideObjects.emplace_back(obj);
                    } while (repeatLine);
                }
            }
            for (const auto &presObj: slide.pictures) {
                if (!presObj.rId.empty()) {
                    insertImage obj;
                    obj.offset = presObj.objectInfo.offsetX / pptInfo.settings.widthCoefficient;
                    obj.startLine = presObj.objectInfo.offsetY / pptInfo.settings.heightCoefficient;
                    obj.length = presObj.objectInfo.objectSizeX / pptInfo.settings.widthCoefficient;
                    obj.rId = presObj.rId;
                    insertImages.emplace_back(obj);
                }
            }
            std::sort(slideObjects.begin(), slideObjects.end());
            slideInsertInfo slideData;
            slideData.slideNum = slideNum++;
            slideData.relations = slide.relations;
            slideData.insertObjects = (slideObjects);
            slideData.insertImages = insertImages;
            slideInsertData.emplace_back(slideData);
        }
    }

    void PptParser::insertSlide(std::vector<insertObject> insertObjects) {
        size_t slideRange = PRESENTATION_WIDTH - 2;
        writer.insertData(std::wstring(L"+").append(slideRange, L'-').append(L"+"), false, false);
        for (int line = 1; line < PRESENTATION_HEIGHT; line++) {
            size_t currentIndex = 0;
            writer.insertData(L"|", true, false);
            for (auto &elem: insertObjects) {
                if (elem.startLine == line) {
                    elem.inProgress = true;
                    insertLineBuffer(elem, &currentIndex);
                } else if (elem.inProgress) {
                    insertLineBuffer(elem, &currentIndex);
                }
            }
            if (slideRange > currentIndex)
                writer.insertData(std::wstring(slideRange - currentIndex, L' '), false, false);
            writer.insertData(L"|", false, false);
        }
        writer.insertData(std::wstring(L"+").append(slideRange, L'-').append(L"+"), true, true);
    }

    void PptParser::insertLineBuffer(insertObject &obj, size_t *currentIndex) {
        writer.insertData(std::wstring(obj.offset - *currentIndex, L' '), false, false);
        auto align = obj.paragraph.front().align;
        auto textSize = obj.paragraph.front().text.length();
        wstring textToInsert;
        if (textSize > obj.length) {
            auto indexLastElement = obj.paragraph.front().text.find_last_of(L' ', obj.length);
            if (indexLastElement == string::npos) {
//                textToInsert = obj.paragraph.front().text.substr(0, obj.length);
            } else {
                textToInsert = obj.paragraph.front().text.substr(0, indexLastElement);
            }
            obj.paragraph.front().text.erase(0, textToInsert.size());
        } else {
            textToInsert = obj.paragraph.front().text;
            obj.paragraph.erase(obj.paragraph.begin());
        }

        switch (align) {
            case l: {
                writer.insertData(textToInsert, false, false);
                if (obj.length - textToInsert.length() != 0)
                    writer.insertData(std::wstring(obj.length - textToInsert.length(), L' '), false, false);
                break;
            }
            case r: {
                if (obj.length - textToInsert.length() != 0)
                    writer.insertData(std::wstring(obj.length - textToInsert.length(), L' '), false, false);
                writer.insertData(textToInsert, false, false);
                break;
            }
            case ctr: {
                if (obj.length - textToInsert.length() != 0)
                    writer.insertData(std::wstring((obj.length - textToInsert.length()) / 2, L' '), false, false);
                writer.insertData(textToInsert, false, false);
                if (obj.length - textToInsert.length() != 0)
                    writer.insertData(std::wstring((obj.length - textToInsert.length()) / 2, L' '), false, false);
                break;
            }

        }
        if (obj.paragraph.empty())
            obj.inProgress = false;
        *currentIndex = writer.getCurrentLength() - 1;
    }

    void PptParser::insertSlideMetadata(slideInsertInfo slideInfo) {
        writer.insertData(L"Slide info: ", true);
        writer.insertData(L"Slide number - " + to_wstring(slideInfo.slideNum));
        if ((options.flags >> 3) & 1)
            for (const auto &kv: slideInfo.relations.notes) {
                auto tmpNote = std::find(pptInfo.notes.begin(), pptInfo.notes.end(), kv.second);
                auto text = tmpNote->text;
                if (!text.empty()) {
                    writer.insertData(L"Slide note: " + text.front().text);

                }
            }
        for (const auto &kv: slideInfo.relations.imageRelationship) {
            auto tmpImage = std::find(slideInfo.insertImages.begin(), slideInfo.insertImages.end(), kv.first);
            writer.insertData(L"Slide image info: " + tmpImage->toString(), false, false);
            if ((options.flags >> 1) & 1) {
                string image = ", saved to path: " + options.pathToDraws + '/' + kv.second;
                writer.insertData(writer.convertString(image));
            }
        }
    }

}

