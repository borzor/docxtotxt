//
// Created by boris on 2/20/22.
//
#include "ParagraphParser.h"
#include "SectionParser.h"
#include <iomanip>
#include <math.h>
#include "DrawingParser.h"

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
                    if (paragraphBuffer.back().length + token.length() / 2 < amountOfCharacters) {
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
                    if (paragraphBuffer.back().length + token.length() < amountOfCharacters) {
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
                case numPr:
                    break;//TODO check for numId and ilvl and somehow link paragraphs into enumeration
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
                    auto drawingParser = Drawing::DrawingParser(textProperty);
                    size_t height, width;
                    drawingParser.parseDrawing();
                    drawingParser.getDrawingSize(height, width);
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
                    break;//TODO tabulation
            }
            textProperty = textProperty->NextSiblingElement();
        }
        free(textProperty);
    }


    void ParagraphParser::writeResult(ofstream &outStream, bool toFile) {
        if (!this->paragraphBuffer.empty()) {
            if (toFile) {

            } else {

            }
            switch (this->justify) {//TODO IDK HOW TO GET WIDTH FOR STREAM SO JUSTIFY WILL WORK
                case left:
                    for (auto &s: paragraphBuffer) {
                        outStream << s.text << '\n';
                    }
                    break;
                case right:
                    for (auto &s: paragraphBuffer) {
                        outStream << setw(amountOfCharacters + strlen(s.text.c_str()) - s.length) << s.text << '\n';
                    }
                    break;
                case center:
                    for (auto &s: paragraphBuffer) {
                        outStream << setw((strlen(s.text.c_str()) + amountOfCharacters / 2 - s.length / 2)) << s.text
                                  << endl;
                    }
                    break;
                case both:
                    for (auto &s: paragraphBuffer) {
                        outStream << s.text << '\n';
                    }
                    break;
                case distribute:
                    break;
            }


        }

    }

    ParagraphParser::ParagraphParser(const size_t amountOfCharacters) {
        this->amountOfCharacters = amountOfCharacters;
        line a;
        paragraphBuffer.push_back(a);
        justify = left;//by default
    }

    void ParagraphParser::setIndentation(XMLElement *element) {
        if (element->Attribute("w:firstLine") != nullptr) {//TWIP_TO_CHARACTER
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

    void ParagraphParser::insertImage(size_t &height, size_t &width) {
        height /= 76200;
        width /= 76200;
        auto leftBorder = (this->amountOfCharacters - width) / 2;
        auto center = height / 2;
        string path = "Here should be image № ";//TODO grep image number and save it if necessary
        for (int i = 0; i < height; i++) {
            line tmp;
            tmp.text.insert(0, leftBorder, ' ');
            if (i != center) {
                tmp.text.insert(leftBorder, width, '#');
            } else {
                tmp.text.insert(leftBorder, (width - path.length()) / 2, '#');
                tmp.text.append(path);
                tmp.text.insert(tmp.text.length(), leftBorder + width - tmp.text.length() + 2, '#');
            }
            tmp.length += amountOfCharacters;
            paragraphBuffer.push_back(tmp);
        }
    }


}