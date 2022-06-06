//
// Created by boris on 2/20/22.
//
#include "../headers/ParagraphParser.h"
#include "../ParserCommons/CommonFunctions.h"
#include <codecvt>
#include <cmath>

namespace docxtotxt {
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
                case outlineLvl: {
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        settings.outline = true;
                }
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
                    if (paragraphProperty->FirstChildElement() != nullptr)
                        setTabulation(paragraphProperty, settings.tab);
                    break;
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
            if(!strcmp(textProperty->Value(), "w:drawing")){
                auto current_element = textProperty->FirstChildElement()->FirstChildElement();
                while (current_element != nullptr) {
                    if (!strcmp(current_element->Value(), "a:graphic")) {
                        if (current_element->FirstChildElement("a:graphicData") != nullptr) {
                            if (current_element->FirstChildElement("a:graphicData")->FirstChildElement("pic:pic") != nullptr) {
                                auto picElement = current_element->FirstChildElement("a:graphicData")->FirstChildElement(
                                        "pic:pic");
                                auto pic = extractPicture(picElement, "pic");
                                insertImage(pic);
                            }
                        }
                    }
                    current_element = current_element->NextSiblingElement();
                }
            } else if (!strcmp(textProperty->Value(), "w:t")){
                auto text = textProperty->GetText();
                if (text == nullptr)
                    paragraphBuffer.append(L" ");
                else
                    paragraphBuffer.append(convertor.from_bytes(text));
            } else if (!strcmp(textProperty->Value(), "w:tab")){
                insertTabulation(settings);
            }
            textProperty = textProperty->NextSiblingElement();
        }
        free(textProperty);
    }

    void ParagraphParser::insertImage(picture pic) {//TODO correct for small images
        auto imageName = docInfo.relations.imageRelationship[pic.rId];
        auto width = pic.objectInfo.objectSizeX / 76200;
        auto height = pic.objectInfo.objectSizeY / 76200;
        auto mediumLine = width - 2;
        auto leftBorder = (docInfo.docWidth - width) / 2;
        wstring path = L"Media file";
        wstring imageInfo = wstring(L"Saved to path: ") + convertor.from_bytes(this->options.pathToDraws) + L'/' +
                            convertor.from_bytes(imageName);
        writer.insertData(std::wstring((docInfo.docWidth - path.length()) / 2, L' ').append(path));
        if ((options.flags >> 1) & 1) {
            writer.insertData(std::wstring((docInfo.docWidth - imageInfo.length()) / 2, L' ').append(imageInfo));
        }
        if (width > 3 || height > 3) {
            writer.insertData(std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—').append(1, L'+'));
            for (int i = 1; i < height - 1; i++) {
                writer.insertData(std::wstring(leftBorder, L' ') + L"|" + std::wstring(mediumLine, L' ') + L"|");
            }
            writer.insertData( std::wstring(leftBorder, L' ').append(1, L'+').append(mediumLine, L'—').append(1, L'+'));
        }
    }

    void ParagraphParser::writeResult() {
        auto justify = settings.justify;
        auto currentSize = paragraphBuffer.size();
        int left = settings.ind.left;
        int right = settings.ind.right;
        int firstLineLeft;
        if (settings.ind.hanging == 0) {
            firstLineLeft = left + settings.ind.firstLine;
        } else {
            firstLineLeft = left == 0 ? 0 : left - settings.ind.hanging;
        }
        bool isFirstLine = true;
        auto before = settings.spacing.before;
        auto after = settings.spacing.after;
        for (int i = 0; i < before; i++) {
            writer.newLine();
        }
        if (!this->paragraphBuffer.empty()) {
            while (currentSize != 0) {
                auto availableBufferInLine = docInfo.docWidth - writer.getCurrentLength();
                if (justify == left)
                    if (isFirstLine) {
                        availableBufferInLine = availableBufferInLine - firstLineLeft - right;
                    } else {
                        availableBufferInLine = availableBufferInLine - left - right;
                    }
                size_t ind = 0;
                if (currentSize > availableBufferInLine) { // 167
                    auto indexLastElement = paragraphBuffer.find_last_of(L' ', availableBufferInLine);
                    if (indexLastElement == string::npos) {
                        writer.newLine();
                        availableBufferInLine = docInfo.docWidth - writer.getCurrentLength();
                        indexLastElement = paragraphBuffer.find_last_of(L' ', availableBufferInLine);
                    }
                    switch (justify) {
                        case paragraphJustify::left: {
                            ind = isFirstLine ? firstLineLeft : left;
                            break;
                        }
                        case paragraphJustify::right: {
                            ind = docInfo.docWidth - indexLastElement;
                            break;
                        }
                        case paragraphJustify::center: {
                            ind = (docInfo.docWidth - indexLastElement) / 2;
                            break;
                        }
                        case both:
                            break;
                        case distribute:
                            break;
                    }
                    writer.insertData(std::wstring(ind, L' ').append(paragraphBuffer.substr(0, indexLastElement)));
                    paragraphBuffer = paragraphBuffer.substr(indexLastElement + 1);
                    isFirstLine = false;
                    currentSize -= indexLastElement;
                } else {
                    switch (justify) {
                        case paragraphJustify::left: {
                            ind = isFirstLine ? firstLineLeft : left;
                            break;
                        }
                        case paragraphJustify::right: {
                            ind = docInfo.docWidth - currentSize;
                            break;
                        }
                        case center: {
                            ind = (docInfo.docWidth - currentSize) / 2;
                            break;
                        }
                        case both:
                            break;
                        case distribute:
                            break;
                    }
                    writer.insertData(std::wstring(ind, L' ').append(paragraphBuffer), false, false);
                    currentSize = 0;
                }
            }
        }
        for (int i = 0; i < after; i++) {
            writer.newLine();
        }
        writer.newLine();
    }

    ParagraphParser::ParagraphParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer)
            : docInfo(docInfo), options(options), writer(writer) {
        settings = docInfo.defaultSettings.paragraph;
    }


    void ParagraphParser::flush() {
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
                auto number = distance(docInfo.relations.hyperlinkRelationship.begin(),
                                       docInfo.relations.hyperlinkRelationship.find(id));//very doubtful
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
                settings = paragraphJustify::left;
            else if (!strcmp(justify.c_str(), "distribute"))
                settings = paragraphJustify::left;
        }
    }

    void ParagraphParser::setIndentation(XMLElement *ind, indentation &settings) {
        if (ind != nullptr) {
            auto left = ind->Attribute("w:left");
            if (left != nullptr) {
                settings.left = atoi(left) / TWIP_TO_CHARACTER;
            }
            auto right = ind->Attribute("w:right");
            if (right != nullptr) {
                settings.right = atoi(right) / TWIP_TO_CHARACTER;
            }
            auto hanging = ind->Attribute("w:hanging");
            if (hanging != nullptr) {
                settings.hanging = atoi(hanging) / TWIP_TO_CHARACTER;
            }
            auto firstLine = ind->Attribute("w:firstLine");
            if (firstLine != nullptr) {
                settings.firstLine = atoi(firstLine) / TWIP_TO_CHARACTER;
            }
        }
    }

    void ParagraphParser::setSpacing(XMLElement *spacing, lineSpacing &settings) {
        if (spacing != nullptr) {
            auto before = spacing->Attribute("w:before");
            if (before != nullptr) {
                settings.before = std::lround(atof(before) / TWIP_TO_CHARACTER);
            }
            auto after = spacing->Attribute("w:after");
            if (after != nullptr) {
                settings.after = std::lround(atof(after) / TWIP_TO_CHARACTER);
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
                if (pos != nullptr) {
                    tmpTab.pos = atoi(pos) / TWIP_TO_CHARACTER;
                }
                settings.emplace(tmpTab);
                tab = tab->NextSiblingElement();
            }
        }
    }

    void ParagraphParser::insertTabulation(paragraphSettings settings) {
        auto tab = settings.tab.front();
        auto pos = tab.pos;
        wchar_t wCharacter = L' ';
        if (tab.character == dot)
            wCharacter = L'.';
        else if (tab.character == heavy)
            wCharacter = L'_';
        else if (tab.character == hyphen)
            wCharacter = L'-';
        auto currentSize = paragraphBuffer.length();
        auto ind = settings.ind.left;
        if (ind != 0)
            currentSize += ind;
        if (pos > currentSize)
            paragraphBuffer.append(pos - currentSize, wCharacter);
        settings.tab.pop();
    }


}