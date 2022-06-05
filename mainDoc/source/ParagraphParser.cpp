//
// Created by boris on 2/20/22.
//
#include "../headers/ParagraphParser.h"
#include "../headers/SectionParser.h"
#include <iomanip>
#include <cmath>


namespace paragraph {
    void Tokenize(const string &str, vector<string> &tokens, const string &delimiters = " ") {
        string::size_type lastPos = str.find_first_not_of(delimiters, 0);
        string::size_type pos = str.find_first_of(delimiters, lastPos);
        while (string::npos != pos || string::npos != lastPos) {
            tokens.push_back(str.substr(lastPos, pos - lastPos));
            lastPos = str.find_first_not_of(delimiters, pos);
            pos = str.find_first_of(delimiters, lastPos);
        }
    }

    void ParagraphParser::addText(const string &text, language language) {
        if (text[0] == ' ') {
            paragraphBuffer.back().text.append(" ");
            paragraphBuffer.back().length += 1;
        }
        switch (language) {
            case ruRU: {
                vector<string> tmp;
                Tokenize(text, tmp);
                for (const auto &token: tmp) {
                    if (paragraphBuffer.back().length + token.length() / 2 < docInfo.docWidth) {
                        if (&token != &tmp.back()) {
                            paragraphBuffer.back().text.append(token).append(" ");
                            paragraphBuffer.back().length += ceil((double) token.length() / 2) + 1;
                        } else {
                            paragraphBuffer.back().text.append(token);
                            paragraphBuffer.back().length += ceil((double) token.length() / 2);
                        }
                    } else {
                        line tmp;
                        tmp.text = token;
                        tmp.text.append(" ");
                        tmp.length = ceil((double) token.length() / 2) + 1;
                        paragraphBuffer.push_back(tmp);
                    }
                }
                if (text.back() == ' ') {
                    paragraphBuffer.back().text.append(" ");
                    paragraphBuffer.back().length += 1;
                }
                break;
            }
            case enUS: {
                vector<string> tmp;
                Tokenize(text, tmp);
                for (const auto &token: tmp) {
                    if (paragraphBuffer.back().length + token.length() < docInfo.docWidth) {
                        if (&token != &tmp.back()) {
                            paragraphBuffer.back().text.append(token).append(" ");
                            paragraphBuffer.back().length += token.length() + 1;
                        } else {
                            paragraphBuffer.back().text.append(token);
                            paragraphBuffer.back().length += token.length();
                        }

                    } else {
                        line tmp;
                        tmp.text = token;
                        tmp.length = token.length();
                        paragraphBuffer.push_back(tmp);
                    }
                }
                if (text.back() == ' ') {
                    paragraphBuffer.back().text.append(" ");
                    paragraphBuffer.back().length += 1;
                }
                break;
            }
        }

    }

    void ParagraphParser::parseParagraph(XMLElement *paragraph) {
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
        XMLElement *paragraphProperty = properties->FirstChildElement();
        while (paragraphProperty != nullptr) {
            switch (paragraphProperties[paragraphProperty->Value()]) {
                case framePr:
                    break;//Defines the paragraph as a text frame, which is a free-standing paragraph similar to a text box
                case ind:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setIndentation(paragraphProperty);
                    break;
                case jc:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setJustify(paragraphProperty->FirstAttribute()->Value());
                    break;
                case keepLines:
                    break;//Specifies that all lines of the paragraph are to be kept on a single page when possible. It is an empty element
                case keepNext:
                    break;//Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next paragraph when possible
                case numPr: {
                    XMLElement *enumProperty = paragraphProperty->FirstChildElement();
                    if (!strcmp(enumProperty->Value(), "w:ilvl")) {
                        if (enumProperty->FirstAttribute() != nullptr) {
                            size_t indentationSize = enumProperty->FirstAttribute()->IntValue();
                            paragraphBuffer.back().text.append(string(indentationSize, ' ')).append(" · ");
                            paragraphBuffer.back().length += indentationSize + 3;
                        }
                    }
                    break;
                }
                case outlineLvl:
                    break;//Specifies the outline level associated with the paragraph
                case pBdr:
                    break;//borders, idk
                case pStyle:
                    break;//TODO add grepping styles from word/styles.xml
                case rPr:
                    break;//styles of text, skip
                case sectPr:
                    break;//idk, should be outside paragraph properties
                case shd:
                    break;//background, skip
                case spacing:
                    break;//spacing between paragraphs, preferably skip
                case tabs:
                    break;//TODO tabulation
                case textAlignment:
                    break;//alignment of characters on each line(if they have different size), skip
            }
            paragraphProperty = paragraphProperty->NextSiblingElement();
        }
        free(paragraphProperty);
    }

    void ParagraphParser::parseTextProperties(XMLElement *properties) {
        XMLElement *textProperty = properties->FirstChildElement();
        language language = ruRU;
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
                    break;//idk, maybe skip
                case rPr://styles of text
                    if (textProperty->FirstChildElement("w:lang") != nullptr)
                        switch (languages[textProperty->FirstChildElement("w:lang")->Value()]) {
                            case ruRU:
                                language = static_cast<enum language>(languages[textProperty->FirstChildElement(
                                        "w:lang")->Value()]);
                                break;
                            case enUS:
                                language = static_cast<enum language> (languages[textProperty->FirstChildElement(
                                        "w:lang")->Value()]);
                                break;
                        }
                    break;
                case softHyphen:
                    break;//never used, optional hyphen character may be added which may appear as a regular hyphen when needed to break the line
                case sym:
                    break;//additional symbol, idk, maybe skip
                case t:
                    if (textProperty->GetText() != nullptr) {
                        addText(textProperty->GetText(), language);
                    } else
                        addText(" ", language);
                    break;
                case tab:
                    paragraphBuffer.back().text.append("    ");
                    paragraphBuffer.back().length += 4;
                    break;
            }
            textProperty = textProperty->NextSiblingElement();
        }
        free(textProperty);
    }


    void ParagraphParser::writeResult() {
        if (!this->paragraphBuffer.empty()) {
            switch (this->justify) {
                case left:
                    for (auto &s: paragraphBuffer) {
                        *options.output << s.text << '\n';
                    }
                    break;
                case right:
                    for (auto &s: paragraphBuffer) {
                        *options.output << setw(docInfo.docWidth + strlen(s.text.c_str()) - s.length) << s.text
                                        << '\n';
                    }
                    break;
                case center:
                    for (auto &s: paragraphBuffer) {
                        *options.output << setw((strlen(s.text.c_str()) + docInfo.docWidth / 2 - s.length / 2))
                                        << s.text
                                        << endl;
                    }
                    break;
                case both:
                    for (auto &s: paragraphBuffer) {
                        *options.output << s.text << '\n';
                    }
                    break;
                case distribute:
                    break;
            }


        }

    }

    ParagraphParser::ParagraphParser(docInfo_t &docInfo, options_t &options)
            : drawingParser(), docInfo(docInfo), options(options) {
        line a;
        paragraphBuffer.push_back(a);
        justify = left;//by default
    }

    void ParagraphParser::setIndentation(XMLElement *element) {
        if (element->Attribute("w:firstLine") != nullptr) {
            auto tmp = atoi(element->Attribute("w:firstLine")) / TWIP_TO_CHARACTER;
            paragraphBuffer.front().text.insert(0, tmp, ' ');
            paragraphBuffer.front().length += tmp;
        } else if (element->Attribute("w:hanging") != nullptr && element->Attribute("w:left") != nullptr) {
            auto tmp = (atoi(element->Attribute("w:hanging")) - atoi(element->Attribute("w:left"))) / TWIP_TO_CHARACTER;
            paragraphBuffer.front().text.insert(0, tmp, ' ');
            paragraphBuffer.front().length += tmp;
        }
    }

    void ParagraphParser::setJustify(const string &justify) {
        if (!strcmp(justify.c_str(), "left"))
            this->justify = left;
        else if (!strcmp(justify.c_str(), "right"))
            this->justify = right;
        else if (!strcmp(justify.c_str(), "center"))
            this->justify = center;
        else if (!strcmp(justify.c_str(), "both"))
            this->justify = both;
        else if (!strcmp(justify.c_str(), "distribute"))
            this->justify = distribute;
    }


    void ParagraphParser::insertImage(size_t &height, size_t &width, const string &imageName) {
        height /= 76200;
        width /= 76200;
        auto leftBorder = (docInfo.docWidth - width) / 2;
        auto center = height / 2 - 1;
        string path = "Media file";
        for (int i = 0; i < height; i++) {
            line tmp;
            tmp.text.insert(0, leftBorder, ' ');
            if (i == center + 1) {
                if ((options.flags >> 1) & 1) {
                    string imageInfo = string("Saved to path: ") + this->options.pathToDraws + '/' + imageName;
                    if (imageInfo.length() > width) {
                        tmp.text.append(imageInfo);
                    } else {
                        tmp.text.insert(leftBorder, (width - imageInfo.length()) / 2, '#');
                        tmp.text.append(imageInfo);
                        tmp.text.insert(tmp.text.length(), leftBorder + width - tmp.text.length(), '#');
                    }
                } else {
                    tmp.text.insert(leftBorder, width, '#');
                }
            } else if (i != center) {
                tmp.text.insert(leftBorder, width, '#');
            } else {
                if (path.length() > width) {
                    tmp.text.append(path);
                } else {
                    tmp.text.insert(leftBorder, (width - path.length()) / 2, '#');
                    tmp.text.append(path);
                    tmp.text.insert(tmp.text.length(), leftBorder + width - tmp.text.length(), '#');
                }
            }
            tmp.length += docInfo.docWidth;
            paragraphBuffer.push_back(tmp);
        }
    }

    void ParagraphParser::flush() {
        options.output->flush();
        paragraphBuffer.clear();
        line a;
        paragraphBuffer.push_back(a);
        justify = left;
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
                auto result = string("{h").append(to_string(number).append("}"));
                paragraphBuffer.back().text.append(result);
                paragraphBuffer.back().length += result.size();
            }
        }
        free(property);
    }

}