//
// Created by boris on 2/9/22.
//

#include <string>
#include <utility>
#include <vector>
#include "parser.h"


namespace prsr {

    void parser::doInit(const string &path) {
        int err;
        struct zip_stat file_info{};
        XMLDocument content;
        zip_t *zip = zip_open(path.c_str(), 0, &err);
        if (zip == nullptr)
            throw runtime_error("Error: Cannot unzip file, error number: " + to_string(err));
        if (zip_stat(zip, "[Content_Types].xml", 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about [Content_Types].xml file");
        char buffer[file_info.size];
        if (zip_fread(zip_fopen(zip, "[Content_Types].xml", 0), &buffer, file_info.size) == -1)
            throw runtime_error("Error: Cannot read [Content_Types].xml file");
        if (content.Parse(buffer) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse [Content_Types].xml file");
        parseContentTypes(&content);
        auto mainPath = this->content_types.find(MAIN_FILE)->second;
        if (mainPath.empty()) {
            std::cout << "Setting name of main document to default\n";
            mainPath = "word/document.xml";
        }

        if (zip_stat(zip, mainPath.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + mainPath + " file");
        char buffer2[file_info.size];
        if (zip_fread(zip_fopen(zip, mainPath.c_str(), 0), &buffer2, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + mainPath + " file");
        if (this->mainDoc.Parse(buffer2) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + mainPath + " file");
        this->isInitialized = true;
        zip_close(zip);
    }

    void parser::checkInit() const {
        if (!isInitialized)
            throw runtime_error("Error: Parser not initialized");
    }

    void parser::parseContentTypes(XMLDocument *xmlDocument) {
        XMLPrinter printer;
        string tmp, tmp2;
        auto *current_element = xmlDocument->RootElement()->FirstChildElement();
        while (current_element != nullptr) {
            current_element->Accept(&printer);
            tmp = printer.CStr();
            auto pos1 = tmp.find('\"');
            auto pos2 = tmp.find('\"', pos1 + 1);
            auto pos3 = tmp.find('\"', pos2 + 1);
            auto pos4 = tmp.find('\"', pos3 + 1);
            tmp2 = tmp.substr(pos1 + 1, pos2 - pos1 - 1);
            if (tmp2.starts_with('/'))
                tmp2.erase(0, 1);
            this->content_types.emplace(tmp.substr(pos3 + 1, pos4 - pos3 - 1), tmp2);
            printer.ClearBuffer();
            current_element = current_element->NextSiblingElement();
        }
        free(current_element);
    }

    parser::parser(string name) :
            name(std::move(name)),
            isInitialized(false) {
        this->paragraphProperties = {{"w:framePr",       framePr},
                                     {"w:ind",           ind},
                                     {"w:jc",            jc},
                                     {"w:keepLines",     keepLines},
                                     {"w:keepNext",      keepNext},
                                     {"w:numPr",         numPr},
                                     {"w:outlineLvl",    outlineLvl},
                                     {"w:pBdr",          pBdr},
                                     {"w:pStyle",        pStyle},
                                     {"w:rPr",           rPr},
                                     {"w:sectPr",        sectPr},
                                     {"w:shd",           shd},
                                     {"w:spacing",       spacing},
                                     {"w:tabs",          tabs},
                                     {"w:textAlignment", textAlignment}};
        this->textProperties = {{"w:br",            br},
                                {"w:cr",            cr},
                                {"w:drawing",       drawing},
                                {"w:noBreakHyphen", noBreakHyphen},
                                {"w:rPr",           rPr},
                                {"w:softHyphen",    softHyphen},
                                {"w:sym",           sym},
                                {"w:t",             t},
                                {"w:tab",           tab},};
    }

    void parser::parseMainDoc() {
        checkInit();
        XMLElement *mainElement = mainDoc.RootElement()->FirstChildElement()->FirstChildElement();
        while (mainElement != nullptr) {
            if (!strcmp(mainElement->Value(), "w:p")) {
                parseParagraph(mainElement);
            } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                parseTable(mainElement);
            } else if (!strcmp(mainElement->Value(), "w:sectPr")) {
                parseSection(mainElement);//TODO maybe should be at start if needed to save page settings
            } else {
                throw runtime_error(string("Unexpected main element: ") + string(mainElement->Value()));
            }
            mainElement = mainElement->NextSiblingElement();
        }
        cout << counter << '\n';
        free(mainElement);

    }

    void parser::parseParagraph(XMLElement *paragraph) {
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
        cout << endl;
        free(property);
    }

    void parser::parseTable(XMLElement *table) {

    }

    void parser::parseSection(XMLElement *section) {

    }

    void parser::parseParagraphProperties(XMLElement *properties) {
        XMLElement *paragraphProperty = properties->FirstChildElement();
        while (paragraphProperty != nullptr) {
            switch (this->paragraphProperties[paragraphProperty->Value()]) {
                case framePr:
                    break;//Defines the paragraph as a text frame, which is a free-standing paragraph similar to a text box
                case ind:
                    break; //TODO Indention
                case jc:
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
                    break;//TODO spacing
                case tabs:
                    break;//TODO tabulation
                case textAlignment:
                    break;//alignment of characters on each line(if they have different size), skip
            }
            paragraphProperty = paragraphProperty->NextSiblingElement();
        }
        free(paragraphProperty);
    }

    void parser::parseTextProperties(XMLElement *properties) {
        XMLElement *textProperty = properties->FirstChildElement();
        while (textProperty != nullptr) {
            switch (this->textProperties[textProperty->Value()]) {
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
                    if (textProperty->GetText() != nullptr)cout << textProperty->GetText();
                    break;//TODO main text, add handler for xml:space
                case tab:
                    break;//TODO tabulation
            }
            textProperty = textProperty->NextSiblingElement();
        }
        free(textProperty);
    }


}
