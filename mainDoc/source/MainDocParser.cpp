//
// Created by boris on 2/9/22.
//

#include <string>
#include <utility>
#include "../headers/MainDocParser.h"
#include "../headers/TableParser.h"


namespace prsr {

    void MainDocParser::doInit() {
        struct zip_stat file_info{};
        zip_stat_init(&file_info);
        XMLDocument content;
        string mainPath, imagesPath;
        int err = 0;
        zip_t *zip = options.input;
        if (zip == nullptr)
            throw runtime_error("Error: Cannot unzip file, error number: " + to_string(err));
        if (zip_stat(zip, "[Content_Types].xml", 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about [Content_Types].xml file");
        char buffer[file_info.size];
        if (zip_fread(zip_fopen(zip, "[Content_Types].xml", 0), &buffer, file_info.size) == -1)
            throw runtime_error("Error: Cannot read [Content_Types].xml file");
        if (content.Parse(buffer, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse [Content_Types].xml file");
        parseContentTypes(&content);
        mainPath = this->content_types.find(MAIN_FILE)->second;
        if (mainPath.empty()) {
            std::cout << "Setting name of main document to default\n";
            mainPath = "word/document.xml";
        }
        if ((options.flags >> 1) & 1) {
            imagesPath = IMAGE_FILE_PATH;
            if (zip_stat(zip, imagesPath.c_str(), 0, &file_info) == -1)
                throw runtime_error("Error: Cannot get info about " + imagesPath + " file");
            char buffer3[file_info.size];
            if (zip_fread(zip_fopen(zip, imagesPath.c_str(), 0), &buffer3, file_info.size) == -1)
                throw runtime_error("Error: Cannot read " + imagesPath + " file");
            if (this->imagesDoc.Parse(buffer3, file_info.size) != tinyxml2::XML_SUCCESS)
                throw runtime_error("Error: Cannot parse " + imagesPath + " file");
            parseImageDoc(&imagesDoc);
            for (const auto &[key, value]: imageRelationship) {
                if (value.ends_with(".png") || value.ends_with(".jpg") || value.ends_with(".jpeg") ||
                    value.ends_with(".gif")) {//TODO
                    if (zip_stat(zip, value.c_str(), ZIP_FL_NODIR, &file_info) == -1)
                        throw runtime_error("Error: Cannot get info about " + value + " file");
                    char tmpImageBuffer[file_info.size];
                    if (zip_fread(zip_fopen(zip, value.c_str(), ZIP_FL_NODIR), &tmpImageBuffer, file_info.size) == -1)
                        throw runtime_error("Error: Cannot read " + value + " file");
                    if (!std::ofstream(string(options.pathToDraws) + "/" + value).write(tmpImageBuffer,
                                                                                        file_info.size)) {
                        throw runtime_error("Error writing file" + value);
                    }
                }
            }
        }
        if (zip_stat(zip, mainPath.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + mainPath + " file");
        char buffer2[file_info.size];
        if (zip_fread(zip_fopen(zip, mainPath.c_str(), 0), &buffer2, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + mainPath + " file");
        if (this->mainDoc.Parse(buffer2, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + mainPath + " file");
        zip_close(zip);
        this->isInitialized = true;
    }

    void MainDocParser::checkInit() const {
        if (!isInitialized)
            throw runtime_error("Error: Parser not initialized");
    }

    void MainDocParser::parseContentTypes(XMLDocument *xmlDocument) {
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

    MainDocParser::MainDocParser(string name, options_t options) :
            name(std::move(name)), options(std::move(options)),
            isInitialized(false) {
    }

    void MainDocParser::parseMainDoc() {
        checkInit();
        XMLElement *mainElement = mainDoc.RootElement()->FirstChildElement()->FirstChildElement();
        XMLElement *section = mainDoc.RootElement()->FirstChildElement()->FirstChildElement("w:sectPr");
        auto sectionParser = section::SectionParser();
        sectionParser.parseSection(section);

        auto paragraphParser = paragraph::ParagraphParser(sectionParser.getDocWidth(), std::move(options),
                                                          imageRelationship);
        auto tableParser = table::TableParser();
        while (mainElement != nullptr) {
            if (!strcmp(mainElement->Value(), "w:p")) {
                paragraphParser.parseParagraph(mainElement);
            } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                tableParser.parseTable(mainElement);
            } else {//TODO think about this
                //throw runtime_error(string("Unexpected main element: ") + string(mainElement->Value()));
            }
            options.output << "test";
            paragraphParser.writeResult();//todo make it more generic, for table+paragraph text
            paragraphParser.clearFields();
            mainElement = mainElement->NextSiblingElement();
        }
        free(mainElement);
        options.output.close();
    }


    void MainDocParser::parseImageDoc(XMLDocument *doc) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        while (mainElement != nullptr) {
            if (!strcmp(mainElement->Value(), "Relationship")) {
                string file = mainElement->Attribute("Target");
                imageRelationship.emplace(mainElement->Attribute("Id"), file.substr(file.find_last_of('/') + 1));
            }
            mainElement = mainElement->NextSiblingElement();
        }
    }


}
