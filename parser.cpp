//
// Created by boris on 2/9/22.
//

#include <string>
#include <vector>
#include "parser.h"


namespace prsr {

    void parser::doInit(const std::string &path) {
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
        if(zip_fread(zip_fopen(zip, mainPath.c_str(), 0), &buffer2, file_info.size) == -1)
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
            if(tmp2.starts_with('/'))
                tmp2.erase(0, 1);
            this->content_types.emplace(tmp.substr(pos3 + 1, pos4 - pos3 - 1), tmp2);
            printer.ClearBuffer();
            current_element = current_element->NextSiblingElement();
        }
        free(current_element);
    }

    parser::parser(const std::string &name) {
        this->name = name;
        this->isInitialized = false;
    }

    void parser::parseMainDoc() {
        checkInit();
        XMLElement *mainElement = mainDoc.RootElement()->FirstChildElement()->FirstChildElement();
        while(mainElement != nullptr){
            if (!strcmp(mainElement->Value(), "w:p")) {
                parseParagraph(mainElement);
            } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                parseTable(mainElement);
            } else if(!strcmp(mainElement->Value(), "w:sectPr")){
                parseSection(mainElement);//TODO maybe should be at start if needed to save page settings
            } else {
                throw runtime_error(string("Unexpected main element: ") + string(mainElement->Value()));
            }
            mainElement = mainElement->NextSiblingElement();
        }

        free(mainElement);
    }

    void parser::parseParagraph(XMLElement *paragraph) {
        XMLElement *paragraphProperties = paragraph->FirstChildElement("w:pPr");
        if (paragraphProperties != nullptr) {
            XMLElement *paragraphProperty = paragraphProperties->FirstChildElement();
            while (paragraphProperty != nullptr) {
                if (!strcmp(paragraphProperty->Value(), "w:framePr")) {
                    //Defines the paragraph as a text frame, which is a free-standing paragraph similar to a text box
                }
                if (!strcmp(paragraphProperty->Value(), "w:ind")) {
                    //TODO Indention
                }
                if (!strcmp(paragraphProperty->Value(), "w:jc")) {
                    //TODO alignment
                }
                if (!strcmp(paragraphProperty->Value(), "w:keepLines")) {//skip
                    //Specifies that all lines of the paragraph are to be kept on a single page when possible. It is an empty element
                }
                if (!strcmp(paragraphProperty->Value(), "w:keepNext")) {//skip
                    //Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next paragraph when possible
                }
                if (!strcmp(paragraphProperty->Value(), "w:numPr")) {//skip
                    //Specifies that the paragraph should be numbered
                }
                if (!strcmp(paragraphProperty->Value(), "w:outlineLvl")) {//idk
                    //Specifies the outline level associated with the paragraph
                }
                if (!strcmp(paragraphProperty->Value(), "w:pBdr")) {
                    //TODO borders
                }
                if (!strcmp(paragraphProperty->Value(), "w:pStyle")) {
                    //TODO add grepping styles from word/styles.xml
                }
                if (!strcmp(paragraphProperty->Value(), "w:rPr")) {//skip
                    //styles of text, skip
                }
                if (!strcmp(paragraphProperty->Value(), "w:sectPr")) {//skip
                    //TODO? idk, should be outside paragraph properties
                }
                if (!strcmp(paragraphProperty->Value(), "w:shd")) {//skip
                    //background, skip
                }
                if (!strcmp(paragraphProperty->Value(), "w:spacing")) {
                    //TODO spacing
                }
                if (!strcmp(paragraphProperty->Value(), "w:tabs")) {
                    //TODO tabulation
                }
                if (!strcmp(paragraphProperty->Value(), "w:textAlignment")) {//skip
                    //alignment of characters on each line(if they have different size), skip
                }
                cout << paragraphProperty->Value() << endl;
                paragraphProperty = paragraphProperty->NextSiblingElement();
            }
            free(paragraphProperty);
        }
        XMLElement *textProperties = paragraph->FirstChildElement("w:r");
        if (textProperties != nullptr) {
            XMLElement *textProperty = textProperties->FirstChildElement();
            while (textProperty != nullptr) {
                if (!strcmp(textProperty->Value(), "w:br")) {
                    //TODO? Line Break
                }
                if (!strcmp(textProperty->Value(), "w:cr")) {
                    //TODO Default line break
                }
                if (!strcmp(textProperty->Value(), "w:drawing")) {
                    //TODO later, images
                }
                if (!strcmp(textProperty->Value(), "w:noBreakHyphen")) {
                    //idk, maybe skip
                }
                if (!strcmp(textProperty->Value(), "w:rPr")) {
                    //style of text, skip
                }
                if (!strcmp(textProperty->Value(), "w:softHyphen")) {
                    //never used, optional hyphen character may be added which may appear as a regular hyphen when needed to break the line
                }
                if (!strcmp(textProperty->Value(), "w:sym")) {
                    //additional symbol, idk, maybe skip
                }
                if (!strcmp(textProperty->Value(), "w:t")) {
                    //TODO main text, add handler for xml:space
                }
                if (!strcmp(textProperty->Value(), "w:tab")) {
                    //TODO tabulation
                }
                cout << textProperty->Value() << endl;
                textProperty = textProperty->NextSiblingElement();
            }
            XMLPrinter printer;
            paragraph->Accept(&printer);
        }
    }

    void parser::parseTable(XMLElement *table) {

    }

    void parser::parseSection(XMLElement *section) {

    }


}
