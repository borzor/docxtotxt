//
// Created by boris on 2/22/22.
//

#include "../headers/DrawingParser.h"

namespace Drawing {
    DrawingParser::DrawingParser(docInfo_t &docInfo, options_t &options) : options(options), docInfo(docInfo) {
    }

    void DrawingParser::parseDrawing(tinyxml2::XMLElement *element) {
        auto current_element = element->FirstChildElement()->FirstChildElement();
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "a:graphic")) {
                if (current_element->FirstChildElement("a:graphicData") != nullptr) {
                    if (current_element->FirstChildElement("a:graphicData")->FirstChildElement("pic:pic") != nullptr) {
                        auto picElement = current_element->FirstChildElement("a:graphicData")->FirstChildElement(
                                "pic:pic")->FirstChildElement();
                        while (picElement != nullptr) {
                            if (!strcmp(picElement->Value(), "pic:nvPicPr")) {
                                //maybe  Specifies non-visual properties of a picture, such as locking properties, name, id and title, and whether the picture is hidden.
                            } else if (!strcmp(picElement->Value(), "pic:blipFill")) {
                                lastImageId = picElement->FirstChildElement()->Attribute("r:embed");
                                //TODO pointer to picture
                            } else if (!strcmp(picElement->Value(), "pic:spPr")) {
                                if (picElement->FirstChildElement("a:xfrm") != nullptr) {
                                    if (picElement->FirstChildElement("a:xfrm")->FirstChildElement("a:ext") !=
                                        nullptr) {
                                        this->height = atoi(picElement->FirstChildElement("a:xfrm")->FirstChildElement(
                                                "a:ext")->Attribute("cy")) / 76200;
                                        this->width = atoi(picElement->FirstChildElement("a:xfrm")->FirstChildElement(
                                                "a:ext")->Attribute("cx")) / 76200;
                                    }
                                }
                            }
                            picElement = picElement->NextSiblingElement();
                        }
                    }
                }
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DrawingParser::insertImage() {//TODO correct for small images
        auto imageName = docInfo.imageRelationship[lastImageId];
        auto mediumLine = width - 2;
        auto leftBorder = (docInfo.docWidth - width) / 2;
        wstring path = L"Media file";
        wstring imageInfo = wstring(L"Saved to path: ") + convertor.from_bytes(this->options.pathToDraws) + L'/' +
                            convertor.from_bytes(imageName);
        addLine(docInfo.resultBuffer);
        docInfo.resultBuffer.buffer.back().append(std::wstring((docInfo.docWidth - path.length()) / 2, L' '));//left border
        docInfo.resultBuffer.buffer.back().append(path);
        if ((options.flags >> 1) & 1) {
            addLine(docInfo.resultBuffer);
            docInfo.resultBuffer.buffer.back().append(
                    std::wstring((docInfo.docWidth - imageInfo.length()) / 2, L' '));//left border
            docInfo.resultBuffer.buffer.back().append(imageInfo);
        }
        addLine(docInfo.resultBuffer);
        if (width > 3 || height > 3) {
            docInfo.resultBuffer.buffer.back().append(
                    std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—')).append(1, L'+');
            for (int i = 1; i < height - 1; i++) {
                addLine(docInfo.resultBuffer);
                docInfo.resultBuffer.buffer.back().append(std::wstring(leftBorder, L' ') + L"|");
                docInfo.resultBuffer.buffer.back().append(std::wstring(mediumLine, L' ') + L"|");
            }
            addLine(docInfo.resultBuffer);
            docInfo.resultBuffer.buffer.back().append(
                    std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—')).append(1, L'+');
        }
        addLine(docInfo.resultBuffer);
    }

        void DrawingParser::flush() {

        }


    }

