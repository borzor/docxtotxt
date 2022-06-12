//
// Created by boris on 2/20/22.
//
#include "../headers/ParagraphParser.h"
#include "../headers/SectionParser.h"
#include <codecvt>

namespace paragraph {
    void ParagraphParser::parseParagraph(XMLElement *paragraph) {
        auto pPr = paragraph->FirstChildElement("w:pPr");
        if (pPr == nullptr) {
            settings = docInfo.styles.defaultStyles.paragraph;
        }
        XMLElement *property = paragraph->FirstChildElement();
        while (property != nullptr) {
            if (!strcmp(property->Value(), "w:pPr")) {
                parseParagraphProperties(property);
            } else if (!strcmp(property->Value(), "w:r")) {
                parseTextProperties(property);
            } else if (!strcmp(property->Value(), "w:hyperlink")) {
                parseHyperlink(property);
            } else {
                //throw runtime_error(string("Unexpected child element: ") + string(property->Value()));
            }
            property = property->NextSiblingElement();
        }
        free(property);
    }

    void ParagraphParser::parseParagraphProperties(XMLElement *properties) {
        auto pStyle = properties->FirstChildElement("w:pStyle");
        if (pStyle == nullptr) {
            settings = docInfo.styles.defaultStyles.paragraph;
        } else {
            settings = docInfo.styles.paragraphStyles[pStyle->Attribute("w:val")];
        }
        XMLElement *paragraphProperty = properties->FirstChildElement();
        while (paragraphProperty != nullptr) {
            switch (paragraphProperties[paragraphProperty->Value()]) {
                case framePr:
                    break;
                case ind:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setIndentation(paragraphProperty, settings.ind);
                    break;
                case jc:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setJustify(paragraphProperty, settings.justify);
                    break;
                case keepLines:
                    break;
                case keepNext:
                    break;
                case numPr: {
                    XMLElement *enumProperty = paragraphProperty->FirstChildElement();
                    if (!strcmp(enumProperty->Value(), "w:ilvl")) {
                        if (enumProperty->FirstAttribute() != nullptr) {
                            size_t indentationSize = enumProperty->FirstAttribute()->IntValue();
                            paragraphBuffer.append(wstring(indentationSize, L' ')).append(L" · ");
                        }
                    }
                    break;
                }
                case outlineLvl:
                    break;
                case pBdr:
                    break;
                case rPr:
                    break;
                case sectPr:
                    break;
                case shd:
                    break;
                case spacing:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setSpacing(paragraphProperty, settings.spacing);
                    break;
                case tabs:
                    break;//TODO another tabulation
                case textAlignment:
                    break;
            }
            paragraphProperty = paragraphProperty->NextSiblingElement();
        }
        free(paragraphProperty);
    }

    void ParagraphParser::parseTextProperties(XMLElement *properties) {
        XMLElement *textProperty = properties->FirstChildElement();
        while (textProperty != nullptr) {
            switch (textProperties[textProperty->Value()]) {
                case br:
                    break;//TODO? Line Break
                case cr:
                    break;//TODO Default line break
                case drawing: {
                    size_t height, width;
                    drawingParser.parseDrawing(textProperty);
                    drawingParser.getDrawingSize(height, width);
                    if ((options.flags >> 1) & 1)
                        insertImage(height, width, docInfo.imageRelationship[drawingParser.getDrawingId()]);
                    else
                        insertImage(height, width);
                    break;
                }
                case noBreakHyphen:
                    break;
                case rPr:
                    break;
                case softHyphen:
                    break;
                case sym:
                    break;
                case t: {
                    auto text = textProperty->GetText();
                    if (text == nullptr)
                        paragraphBuffer.append(L" ");
                    else
                        paragraphBuffer.append(convertor.from_bytes(text));
                    break;
                }
                case tab:
                    paragraphBuffer.append(L"    ");
                    break;
            }
            textProperty = textProperty->NextSiblingElement();
        }
        free(textProperty);
    }


    void ParagraphParser::writeResult() {
        auto currentSize = paragraphBuffer.size();
        size_t left = settings.ind.left / TWIP_TO_CHARACTER;
        size_t right = settings.ind.right / TWIP_TO_CHARACTER;
        size_t firstLineLeft;
        if (settings.ind.hanging == 0) {
            firstLineLeft = left + settings.ind.firstLine / TWIP_TO_CHARACTER;
        } else {
            firstLineLeft = left == 0 ? 0 : left - settings.ind.hanging / TWIP_TO_CHARACTER;
        }
        bool isFirstLine = true;
        auto before = settings.spacing.before / TWIP_TO_CHARACTER_HEIGHT;
        auto after = settings.spacing.after / TWIP_TO_CHARACTER_HEIGHT;
        for (int i = 0; i < before; i++) {
            docInfo.docBuffer.emplace_back();
            docInfo.pointer++;
        }
        if (!this->paragraphBuffer.empty()) {
            while (currentSize != 0) {
                auto currentLine = docInfo.pointer;
                if (docInfo.docBuffer[currentLine].second == -1)
                    docInfo.docBuffer.emplace_back();
                auto availableBufferInLine = docInfo.docWidth - docInfo.docBuffer[currentLine].second;
                if (isFirstLine) {
                    availableBufferInLine = availableBufferInLine - firstLineLeft - right;
                } else {
                    availableBufferInLine = availableBufferInLine - left - right;
                }
                if (currentSize > availableBufferInLine) {
                    auto indexLastElement = paragraphBuffer.find_last_of(L' ', availableBufferInLine);
                    if (indexLastElement == string::npos) {
                        if (isFirstLine) {
                            availableBufferInLine = docInfo.docWidth - firstLineLeft - right;
                        } else {
                            availableBufferInLine = docInfo.docWidth - left - right;
                        }
                        indexLastElement = paragraphBuffer.find_last_of(L' ', availableBufferInLine);
                        docInfo.docBuffer.emplace_back();
                        docInfo.pointer++;
                    }
                    switch (settings.justify) {
                        case paragraphJustify::left: {
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer.substr(0, indexLastElement));
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second += indexLastElement;
                            currentSize -= indexLastElement;
                            paragraphBuffer = paragraphBuffer.substr(indexLastElement);
                            docInfo.docBuffer.emplace_back();
                            docInfo.pointer++;
                            isFirstLine = false;
                            break;
                        }
                        case paragraphJustify::right: {
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append(docInfo.docWidth - indexLastElement, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer.substr(0, indexLastElement));
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second += indexLastElement;
                            currentSize -= indexLastElement;
                            paragraphBuffer = paragraphBuffer.substr(indexLastElement);
                            docInfo.docBuffer.emplace_back();
                            docInfo.pointer++;
                            isFirstLine = false;
                            break;
                        }
                        case paragraphJustify::center: {
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append((docInfo.docWidth - indexLastElement) / 2, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer.substr(0, indexLastElement));
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second += indexLastElement;
                            currentSize -= indexLastElement;
                            paragraphBuffer = paragraphBuffer.substr(indexLastElement);
                            docInfo.docBuffer.emplace_back();
                            docInfo.pointer++;
                            isFirstLine = false;
                            break;
                        }
                        case both:
                            break;
                        case distribute:
                            break;
                    }
                } else {
                    switch (settings.justify) {
                        case paragraphJustify::left: {
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer);
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second +=
                                    paragraphBuffer.length() + (isFirstLine ? firstLineLeft + right : left + right);
                            currentSize = 0;
                            break;
                        }
                        case paragraphJustify::right: {
                            auto insertingSize =
                                    paragraphBuffer.length() + (isFirstLine ? firstLineLeft + right : left + right);
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append(docInfo.docWidth - insertingSize, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer);
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second += insertingSize;
                            currentSize = 0;
                            break;
                        }
                        case paragraphJustify::center: {
                            auto insertingSize =
                                    paragraphBuffer.length() + (isFirstLine ? firstLineLeft + right : left + right);
                            docInfo.docBuffer.back().first.append(isFirstLine ? firstLineLeft : left, L' ');
                            docInfo.docBuffer.back().first.append((docInfo.docWidth - insertingSize) / 2, L' ');
                            docInfo.docBuffer.back().first.append(paragraphBuffer);
                            docInfo.docBuffer.back().first.append(right, L' ');
                            docInfo.docBuffer.back().second += insertingSize;
                            currentSize = 0;
                            break;
                        }
                        case both:
                            break;
                        case distribute:
                            break;
                    }
                }
            }
        }
        for (int i = 0; i < after; i++) {
            docInfo.docBuffer.emplace_back();
            docInfo.pointer++;
        }
        docInfo.docBuffer.emplace_back();
        docInfo.pointer++;
    }

    ParagraphParser::ParagraphParser(docInfo_t &docInfo, options_t &options)
            : drawingParser(), docInfo(docInfo), options(options) {
        settings = docInfo.defaultSettings.paragraph;
    }

    void ParagraphParser::insertImage(size_t &height, size_t &width, const string &imageName) {
        height /= 76200;
        width /= 76200;
        auto leftBorder = (docInfo.docWidth - width) / 2;
        auto center = height / 2 - 1;
        wstring path = L"Media file";
        for (int i = 0; i < height; i++) {
            wstring tmp;
            tmp.insert(0, leftBorder, ' ');
            if (i == center + 1) {
                if ((options.flags >> 1) & 1) {
                    wstring imageInfo =
                            wstring(L"Saved to path: ") + convertor.from_bytes(this->options.pathToDraws) + L'/' +
                            convertor.from_bytes(imageName);
                    if (imageInfo.length() > width) {
                        tmp.append(imageInfo);
                    } else {
                        tmp.insert(leftBorder, (width - imageInfo.length()) / 2, '#');
                        tmp.append(imageInfo);
                        tmp.insert(tmp.length(), leftBorder + width - tmp.length(), '#');
                    }
                } else {
                    tmp.insert(leftBorder, width, '#');
                }
            } else if (i != center) {
                tmp.insert(leftBorder, width, '#');
            } else {
                if (path.length() > width) {
                    tmp.append(path);
                } else {
                    tmp.insert(leftBorder, (width - path.length()) / 2, '#');
                    tmp.append(path);
                    tmp.insert(tmp.length(), leftBorder + width - tmp.length(), '#');
                }
            }
            paragraphBuffer.append(tmp);
        }
    }

    void ParagraphParser::flush() {//TODO check correct flush
        options.output->flush();
        paragraphBuffer.clear();
        settings = docInfo.defaultSettings.paragraph;
    }

    void ParagraphParser::parseHyperlink(XMLElement *properties) {
        XMLElement *property = properties->FirstChildElement();
        while (property != nullptr) {
            if (!strcmp(property->Value(), "w:r")) {
                parseTextProperties(property);
            }
            property = property->NextSiblingElement();
        }
        if ((options.flags >> 2) & 1) {
            if (properties->Attribute("r:id") != nullptr) {
                auto id = properties->Attribute("r:id");
                auto number = distance(docInfo.hyperlinkRelationship.begin(),
                                       docInfo.hyperlinkRelationship.find(id));//very doubtful
                auto result = wstring(L"{h").append(to_wstring(number).append(L"}"));
                paragraphBuffer.append(result);
            }
        }
        free(property);
    }

    wstring ParagraphParser::getResult() {
        return this->paragraphBuffer;
    }

    void ParagraphParser::setJustify(XMLElement *jc, paragraphJustify &settings) {
        if (jc != nullptr) {
            string justify = jc->Attribute("w:val");
            if (!strcmp(justify.c_str(), "left"))
                settings = paragraphJustify::left;
            else if (!strcmp(justify.c_str(), "right"))
                settings = paragraphJustify::right;
            else if (!strcmp(justify.c_str(), "center"))
                settings = paragraphJustify::center;
            else if (!strcmp(justify.c_str(), "both"))
                settings = paragraphJustify::both;
            else if (!strcmp(justify.c_str(), "distribute"))
                settings = paragraphJustify::distribute;
        }
    }

    void ParagraphParser::setIndentation(XMLElement *ind, indentation &settings) {
        if (ind != nullptr) {
            auto left = ind->Attribute("w:left");
            if (left != nullptr) {
                settings.left = atoi(left);
            }
            auto right = ind->Attribute("w:right");
            if (right != nullptr) {
                settings.right = atoi(right);
            }
            auto hanging = ind->Attribute("w:hanging");
            if (hanging != nullptr) {
                settings.hanging = atoi(hanging);
            }
            auto firstLine = ind->Attribute("w:firstLine");
            if (firstLine != nullptr) {
                settings.firstLine = atoi(firstLine);
            }
        }
    }

    void ParagraphParser::setSpacing(XMLElement *spacing, lineSpacing &settings) {
        if (spacing != nullptr) {
            auto before = spacing->Attribute("w:before");
            if (before != nullptr) {
                settings.before = atoi(before);
            }
            auto after = spacing->Attribute("w:after");
            if (after != nullptr) {
                settings.after = atoi(after);
            }
        }
    }

    void ParagraphParser::setTabulation(XMLElement *tabs, queue<tabulation> &settings) {
        if (tabs != nullptr) {
            auto *tab = tabs->FirstChildElement();
            while (tab != nullptr) {
                tabulation tmpTab;
                tmpTab.character = none;
                auto leader = tab->Attribute("w:leader");
                if (leader != nullptr) {
                    if (!strcmp(leader, "dot")) {
                        tmpTab.character = dot;
                    } else if (!strcmp(leader, "heavy")) {
                        tmpTab.character = heavy;
                    } else if (!strcmp(leader, "hyphen")) {
                        tmpTab.character = hyphen;
                    } else if (!strcmp(leader, "middleDot")) {
                        tmpTab.character = middleDot;
                    } else if (!strcmp(leader, "none")) {
                        tmpTab.character = none;
                    } else if (!strcmp(leader, "underScore")) {
                        tmpTab.character = underscore;
                    }
                }
                auto pos = tab->Attribute("w:pos");
                if(pos != nullptr){
                    tmpTab.pos = atoi(pos);
                }
                settings.emplace(tmpTab);
                tab = tab->NextSiblingElement();
            }
        }
    }


}