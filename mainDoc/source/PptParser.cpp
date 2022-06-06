//
// Created by borzor on 12/12/22.
//

#include "../headers/PptParser.h"

namespace ppt {
    PptParser::PptParser(pptInfo_t &pptInfo, options_t &options) : pptInfo(pptInfo), options(options) {

    }

    void PptParser::parseSlides() {
        prepareSlides();
        pptInfo.documentData.resultBuffer.newLine();
        for (auto &obj: slideInsertData) {
            insertSlideMetaData(obj);
            pptInfo.documentData.resultBuffer.newLine();
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
                    obj.length = presObj.objectInfo.objectSizeX/ pptInfo.settings.widthCoefficient;
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
                                    text.text.append(txt.text.substr(0, table.gridColSize[colNum] - 1));
                                    txt.text = txt.text.substr(table.gridColSize[colNum] - 1);
                                    text.text.append(1, ' ');
                                    repeatLine = true;
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
                        obj.startLine = table.objectInfo.offsetY  / pptInfo.settings.heightCoefficient +counter;
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
                    obj.length = presObj.objectInfo.objectSizeX/ pptInfo.settings.widthCoefficient;
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
        pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(1, L'+').append(slideRange, L'—')).append(1,
                                                                                                                      L'+');
        for (int line = 1; line < PRESENTATION_HEIGHT; line++) {
            pptInfo.documentData.resultBuffer.newLine();
            size_t currentIndex = 0;
            pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(1, L'|'));
            for (auto &elem: insertObjects) {
                if (elem.startLine == line) {
                    elem.inProgress = true;
                    insertLineBuffer(elem, &currentIndex);
                } else if (elem.inProgress) {
                    insertLineBuffer(elem, &currentIndex);
                }
            }
            if (slideRange > currentIndex)
                pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(slideRange - currentIndex, L' '));
            pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(1, L'|'));
        }
        pptInfo.documentData.resultBuffer.newLine();
        pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(1, L'+').append(slideRange, L'—')).append(1,
                                                                                                                      L'+');
        pptInfo.documentData.resultBuffer.newLine();
    }

    void PptParser::insertLineBuffer(insertObject &obj, size_t *currentIndex) {
        pptInfo.documentData.resultBuffer.buffer.back().append(std::wstring(obj.offset - *currentIndex, L' '));
        auto align = obj.paragraph.front().align;
        auto textSize = obj.paragraph.front().text.length();
        wstring textToInsert;
        if (textSize > obj.length) {
            auto indexLastElement = obj.paragraph.front().text.find_last_of(L' ', obj.length);
            if (indexLastElement == string::npos) {
                textToInsert = obj.paragraph.front().text.substr(0, obj.length);
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
                pptInfo.documentData.resultBuffer.buffer.back().append(textToInsert);
                if (obj.length - textToInsert.length() != 0)
                    pptInfo.documentData.resultBuffer.buffer.back().append(
                            std::wstring(obj.length - textToInsert.length(), L' '));
                break;
            }
            case r: {
                if (obj.length - textToInsert.length() != 0)
                    pptInfo.documentData.resultBuffer.buffer.back().append(
                            std::wstring(obj.length - textToInsert.length(), L' '));
                pptInfo.documentData.resultBuffer.buffer.back().append(textToInsert);
                break;
            }
            case ctr: {
                if (obj.length - textToInsert.length() != 0)
                    pptInfo.documentData.resultBuffer.buffer.back().append(
                            std::wstring((obj.length - textToInsert.length()) / 2, L' '));
                pptInfo.documentData.resultBuffer.buffer.back().append(textToInsert);
                if (obj.length - textToInsert.length() != 0)
                    pptInfo.documentData.resultBuffer.buffer.back().append(
                            std::wstring((obj.length - textToInsert.length()) / 2, L' '));
                break;
            }

        }
        if (obj.paragraph.empty())
            obj.inProgress = false;
        *currentIndex = pptInfo.documentData.resultBuffer.buffer.back().length() - 1;
    }

    void PptParser::insertSlideMetaData(slideInsertInfo slideInfo) {
        pptInfo.documentData.resultBuffer.newLine();
        pptInfo.documentData.resultBuffer.buffer.back().append(L"Slide info: ");
        pptInfo.documentData.resultBuffer.buffer.back().append(L"Slide number - ").append(to_wstring(slideInfo.slideNum));
        if ((options.flags >> 3) & 1)
            for (const auto &kv: slideInfo.relations.notes) {
                auto tmpNote = std::find(pptInfo.notes.begin(), pptInfo.notes.end(), kv.second);
                auto text = tmpNote->text;
                if (!text.empty()) {
                    pptInfo.documentData.resultBuffer.newLine();
                    pptInfo.documentData.resultBuffer.buffer.back().append(L"Slide note: ").append(text.front().text);

                }
            }
        for (const auto &kv: slideInfo.relations.imageRelationship) {
            auto tmpImage = std::find(slideInfo.insertImages.begin(), slideInfo.insertImages.end(), kv.first);
            pptInfo.documentData.resultBuffer.newLine();
            pptInfo.documentData.resultBuffer.buffer.back().append(L"Slide image info: ").append(tmpImage->toString());
            if ((options.flags >> 1) & 1) {
                string image = ", saved to path: " + options.pathToDraws + '/' + kv.second;
                pptInfo.documentData.resultBuffer.buffer.back().append(convertor.from_bytes(image));
            }
        }
    }

}

