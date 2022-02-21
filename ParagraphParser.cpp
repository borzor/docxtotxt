//
// Created by boris on 2/20/22.
//
#include "ParagraphParser.h"
#include <iomanip>
#include <math.h>

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
            case RU: {
                vector<string> tmp;
                Tokenize(text, tmp);
                for (const auto &token: tmp) {
                    if (paragraphBuffer.back().length + token.length() / 2 < SIZE_OF_PAGE) {
                        if (&token != &tmp.back()) {
                            paragraphBuffer.back().text.append(token).append(" ");
                            paragraphBuffer.back().length += ceil((double)token.length() / 2) + 1;
                        } else {
                            paragraphBuffer.back().text.append(token);
                            paragraphBuffer.back().length += ceil((double)token.length() / 2);
                        }
                    } else {
                        line tmp;
                        tmp.text = token;
                        tmp.length = token.length() / 2;
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
                    if (paragraphBuffer.back().length + token.length() < SIZE_OF_PAGE) {
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
                    break; //TODO Indention
                case jc:
                    if (paragraphProperty->FirstAttribute() != nullptr)
                        setJustify(paragraphProperty->FirstAttribute()->Value());
                    break; //TODO alignment
                case keepLines:
                    break;//Specifies that all lines of the paragraph are to be kept on a single page when possible. It is an empty element
                case keepNext:
                    break;//Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next paragraph when possible
                case numPr:
                    break;//Specifies that the paragraph should be numbered
                case outlineLvl:
                    break;//Specifies the outline level associated with the paragraph
                case pBdr:
                    break;//TODO borders
                case pStyle:
                    break;//TODO add grepping styles from word/styles.xml
                case rPr:
                    break;//styles of text, skip
                case sectPr:
                    break;//TODO? idk, should be outside paragraph properties
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
        language language = RU;
        while (textProperty != nullptr) {
            switch (textProperties[textProperty->Value()]) {
                case br:
                    break;//TODO? Line Break
                case cr:
                    break;//TODO Default line break
                case drawing:
                    break;//TODO later, images
                case noBreakHyphen:
                    break;//idk, maybe skip
                case rPr:
                    //TODO GET <w:lang w:val="en-US"/> TO SET LANGUAGE TO ENG
                    break;//style of text, skip
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
                        outStream << setw(SIZE_OF_PAGE + strlen(s.text.c_str()) - s.length)<< s.text << '\n';
                    }
                    break;
                case center:
                    for (auto &s: paragraphBuffer) {
                        outStream << setw((strlen(s.text.c_str()) + SIZE_OF_PAGE / 2 - s.length / 2)) << s.text << endl;
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

    ParagraphParser::ParagraphParser() {
        line a;
        paragraphBuffer.push_back(a);
        justify = left;//by default
    }


}