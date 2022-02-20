//
// Created by boris on 2/20/22.
//

#include "ParagraphParser.h"
#include <sys/ioctl.h> //ioctl() and TIOCGWINSZ
#include <unistd.h> // for STDOUT_FILENO
namespace paragraph {

    void ParagraphParser::addText(const string &text) {
        this->paragraphBuffer.push_back(text);
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
                    break;//style of text, skip
                case softHyphen:
                    break;//never used, optional hyphen character may be added which may appear as a regular hyphen when needed to break the line
                case sym:
                    break;//additional symbol, idk, maybe skip
                case t:
                    if (textProperty->GetText() != nullptr) {
                        addText(textProperty->GetText());
                    } else
                        addText(" ");
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
            if(toFile){

            } else{

            }
            switch (this->justify) {//TODO IDK HOW TO GET WIDTH FOR STREAM SO JUSTIFY WILL WORK
                case left:
                    outStream.setf(ios_base::left);
                    break;
                case right:
                    outStream.setf(ios_base::right);
                    break;
                case center:
                    break;
                case both:
                    break;
                case distribute:
                    break;
            }
            for (const auto &e: paragraphBuffer) {
                outStream << e;
            }
            outStream << '\n';
            outStream.unsetf(ios_base::adjustfield);
        }

    }


}