//
// Created by boris on 2/9/22.
//

#include <string>
#include <locale>

#include "../headers/MainDocParser.h"
#include "../headers/TableParser.h"


namespace prsr {
    void MainDocParser::doInit() {
        string mainFileNameProperty, defaultMainPath, mainFileStyleProperty, defaultStylePath;
        struct zip_stat file_info{};
        zip_stat_init(&file_info);
        XMLDocument content, stylesDoc, relationsDoc;
        int err = 0;
        if (options.input == nullptr)
            throw runtime_error("Error: Cannot unzip file, error number: " + to_string(err));
        if (zip_stat(options.input, "[Content_Types].xml", 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about [Content_Types].xml file");
        char buffer[file_info.size];
        auto ContentTypesZip = zip_fopen(options.input, "[Content_Types].xml", 0);
        if (zip_fread(ContentTypesZip, &buffer, file_info.size) == -1)
            throw runtime_error("Error: Cannot read [Content_Types].xml file");
        zip_fclose(ContentTypesZip);
        if (content.Parse(buffer, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse [Content_Types].xml file");
        parseContentTypes(&content);
        switch (options.docType) {
            case pptx: {
                mainFileNameProperty = PPT_MAIN_FILE;
                defaultMainPath = "ppt/presentation.xml";
                mainFileStyleProperty = PPT_STYLES_FILE;
                defaultStylePath = "ppt/tableStyles.xml";
                parsePresentationSpecificData();
                break;
            }
            case docx: {
                mainFileNameProperty = DOC_MAIN_FILE;
                defaultMainPath = "word/document.xml";
                mainFileStyleProperty = DOC_STYLES_FILE;
                defaultStylePath = "word/styles.xml";
                parseDocSpecificData(&relationsDoc);
                break;
            }
            case xlsx: {
                mainFileNameProperty = DOC_MAIN_FILE;
                break;
            }
        }
        std::string mainPath; //this->content_types.find(mainFileNameProperty)->second;
        for (const auto& kv: content_types){
            if(kv.second == mainFileNameProperty){
                mainPath = kv.first;
            }
        }
        if (mainPath.empty()) {
            std::cout << "Setting name of main document to default\n";
            mainPath = defaultMainPath;
        }
        if (zip_stat(options.input, mainPath.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + mainPath + " file");
        char buffer2[file_info.size];
        auto zipMainDoc = zip_fopen(options.input, mainPath.c_str(), 0);
        if (zip_fread(zipMainDoc, &buffer2, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + mainPath + " file");
        zip_fclose(zipMainDoc);
        if (this->mainDoc.Parse(buffer2, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + mainPath + " file");
        string stylesPath;// = this->content_types.find(mainFileStyleProperty)->second;
        for (const auto& kv: content_types){
            if(kv.second == mainFileStyleProperty){
                stylesPath = kv.first;
            }
        }

        if (stylesPath.empty()) {
            std::cout << "Setting name of styles document to default\n";
            stylesPath = defaultStylePath;
        }
        if (zip_stat(options.input, stylesPath.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + stylesPath + " file");
        char buffer3[file_info.size];
        auto zipStylesDoc = zip_fopen(options.input, stylesPath.c_str(), 0);
        if (zip_fread(zipStylesDoc, &buffer3, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + stylesPath + " file");
        zip_fclose(zipStylesDoc);
        if (stylesDoc.Parse(buffer3, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + stylesPath + " file");
        parseStyles(&stylesDoc);
        zip_close(options.input);
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
            if (tmp2.rfind('/', 0) == 0)
                tmp2.erase(0, 1);
            this->content_types.emplace(tmp2, tmp.substr(pos3 + 1, pos4 - pos3 - 1));
            printer.ClearBuffer();
            current_element = current_element->NextSiblingElement();
        }
        free(current_element);
    }

    MainDocParser::MainDocParser(options_t &options, docInfo_t &docInfo, pptInfo_t &pptInfo) :
            options(options),
            isInitialized(false), docInfo(docInfo), pptInfo(pptInfo) {
        docInfo.docBuffer.emplace_back();
    }

    void MainDocParser::parseMainDoc() {
        checkInit();
        switch (options.docType) {
            case pptx: {
                break;
            }
            case docx: {
                XMLElement *mainElement = mainDoc.RootElement()->FirstChildElement()->FirstChildElement();
                XMLElement *section = mainDoc.RootElement()->FirstChildElement()->FirstChildElement("w:sectPr");
                auto sectionParser = section::SectionParser(docInfo);
                sectionParser.parseSection(section);
                paragraph::ParagraphParser paragraphParser(docInfo, options);
                auto tableParser = table::TableParser(docInfo, options);
                while (mainElement != nullptr) {
                    if (!strcmp(mainElement->Value(), "w:p")) {
                        paragraphParser.parseParagraph(mainElement);
                        paragraphParser.writeResult();
                        paragraphParser.flush();
                    } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                        tableParser.parseTable(mainElement);
                    } else {
                        //throw runtime_error(string("Unexpected main element: ") + string(mainElement->Value()));
                    }
                    mainElement = mainElement->NextSiblingElement();
                }
                for (const auto &kv: docInfo.docBuffer) {
                    *options.output << kv.first << std::endl;
                }
                if ((options.flags >> 2) & 1) {
                    insertHyperlinks();
                }
                free(mainElement);
                break;
            }
            case xlsx: {
                break;
            }
        }
    }


    void MainDocParser::parseRelationships(XMLDocument *doc) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        while (mainElement != nullptr) {
            if (!strcmp(mainElement->Value(), "Relationship")) {
                auto type = mainElement->Attribute("Type");
                auto id = mainElement->Attribute("Id");
                string target = mainElement->Attribute("Target");
                if (ends_with(type, "image")) {
                    docInfo.imageRelationship.emplace(id, target.substr(target.find_last_of('/') + 1));
                } else if (ends_with(type, "hyperlink")) {
                    docInfo.hyperlinkRelationship.emplace(id, target);
                }
            }
            mainElement = mainElement->NextSiblingElement();
        }
    }

    void MainDocParser::insertHyperlinks() {
        for (const auto &kv: docInfo.hyperlinkRelationship) {
            auto number = distance(docInfo.hyperlinkRelationship.begin(),
                                   docInfo.hyperlinkRelationship.find(kv.first));//very doubtful
            auto result = wstring(L"{h").append(
                    to_wstring(number).append(L"} -  ").append(convertor.from_bytes(kv.second)));
            *options.output << result << '\n';
            options.output->flush();
        }
    }

    void MainDocParser::saveImages() {
        struct zip_stat file_info{};
        zip_stat_init(&file_info);
        for (const auto &kv: docInfo.imageRelationship) {
            if (zip_stat(options.input, kv.second.c_str(), ZIP_FL_NODIR, &file_info) == -1)
                throw runtime_error("Error: Cannot get info about " + kv.second + " file");
            char tmpImageBuffer[file_info.size];
            auto currentImage = zip_fopen(options.input, kv.second.c_str(), ZIP_FL_NODIR);
            if (zip_fread(currentImage, &tmpImageBuffer, file_info.size) == -1)
                throw runtime_error("Error: Cannot read " + kv.second + " file");
            zip_fclose(currentImage);
            if (!std::ofstream(string(options.pathToDraws) + "/" + kv.second).write(tmpImageBuffer, file_info.size)) {
                throw runtime_error("Error writing file" + kv.second);
            }
        }
    }

    void MainDocParser::parseStyles(XMLDocument *doc) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        switch (options.docType) {
            case pptx:
                break;
            case docx:{
                while (mainElement != nullptr) {
                    if (!strcmp(mainElement->Value(), "w:style")) {
                        addStyle(mainElement);
                    } else if (!strcmp(mainElement->Value(), "w:docDefaults")) {
                        setDefaultSettings(mainElement);
                    } else if (!strcmp(mainElement->Value(), "w:latentStyles")) {
                        //skip
                    }
                    mainElement = mainElement->NextSiblingElement();
                }
                break;
            }
            case xlsx:
                break;
        }


    }

    void MainDocParser::addStyle(XMLElement *element) {
        string type = element->Attribute("w:type");
        string styleId = element->Attribute("w:styleId");
        bool def = false;
        if (element->Attribute("w:default") != nullptr) {
            def = true;
        }
        if (!strcmp(type.c_str(), "paragraph")) {
            paragraphSettings styleSettings = {};
            auto pPr = element->FirstChildElement("w:pPr");
            if (pPr != nullptr) {
                paragraph::ParagraphParser::setIndentation(pPr->FirstChildElement("w:ind"), styleSettings.ind);
                paragraph::ParagraphParser::setJustify(pPr->FirstChildElement("w:jc"), styleSettings.justify);
                paragraph::ParagraphParser::setSpacing(pPr->FirstChildElement("w:spacing"), styleSettings.spacing);
                paragraph::ParagraphParser::setTabulation(pPr->FirstChildElement("w:tabs"), styleSettings.tab);
            }
            if (def) {
                docInfo.styles.defaultStyles.paragraph = styleSettings;
            } else {
                docInfo.styles.paragraphStyles.emplace(styleId, styleSettings);
            }
        } else if (!strcmp(type.c_str(), "table")) {
            tableProperties styleSettings = {};
            auto tblPr = element->FirstChildElement("w:tblPr");
            if (tblPr != nullptr) {
                table::TableParser::setJustify(tblPr->FirstChildElement("w:tblInd"), styleSettings.justify);
                table::TableParser::setIndentation(tblPr->FirstChildElement("w:jc"), styleSettings.ind);
                table::TableParser::setFloatingSettings(tblPr->FirstChildElement("w:tblpPr"), styleSettings.floatTable);
            }
            if (def) {
                docInfo.styles.defaultStyles.table = styleSettings;
            } else {
                docInfo.styles.tableStyles.emplace(styleId, styleSettings);
            }
        } else if (!strcmp(type.c_str(), "character")) {

        } else if (!strcmp(type.c_str(), "numbering")) {

        }

    }

    void MainDocParser::setDefaultSettings(XMLElement *element) {
        auto *docDefault = element->FirstChildElement();
        docInfo.defaultSettings.paragraph.justify = paragraphJustify::left;
        while (docDefault != nullptr) {
            if (!strcmp(docDefault->Value(), "w:rPrDefault")) {//default text properties

            } else if (!strcmp(docDefault->Value(), "w:pPrDefault")) {//default paragraph properties
                auto pPr = docDefault->FirstChildElement("w:pPr");
                if (pPr != nullptr) {
                    paragraph::ParagraphParser::setIndentation(pPr->FirstChildElement("w:ind"),
                                                               docInfo.defaultSettings.paragraph.ind);
                    paragraph::ParagraphParser::setJustify(pPr->FirstChildElement("w:jc"),
                                                           docInfo.defaultSettings.paragraph.justify);
                }
            }
            docDefault = docDefault->NextSiblingElement();
        }
    }
    void MainDocParser::parsePresentationSpecificData(){
        struct zip_stat file_info{};
        string slideContentType = PPT_SLIDE_CONTENT_TYPE;
        for (const auto &kv: this->content_types){
            if(starts_with(kv.second, slideContentType)){
                auto fileName = kv.first;
                XMLDocument document;
                if (zip_stat(options.input, fileName.c_str(), 0, &file_info) == -1)
                    throw runtime_error("Error: Cannot get info about " + fileName + " file");
                char buffer4[file_info.size];
                auto imagesFile = zip_fopen(options.input, fileName.c_str(), 0);
                if (zip_fread(imagesFile, &buffer4, file_info.size) == -1)
                    throw runtime_error("Error: Cannot read " + fileName + " file");
                zip_fclose(imagesFile);
                if (document.Parse(buffer4, file_info.size) != tinyxml2::XML_SUCCESS)
                    throw runtime_error("Error: Cannot parse " + fileName + " file");
                pptInfo.slides.emplace(fileName, &document);
            }
        }

    }

    void MainDocParser::parseDocSpecificData(XMLDocument *document){
        struct zip_stat file_info{};
        string imagesPath = DOC_IMAGE_FILE_PATH;
        if (zip_stat(options.input, imagesPath.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + imagesPath + " file");
        char buffer4[file_info.size];
        auto imagesFile = zip_fopen(options.input, imagesPath.c_str(), 0);
        if (zip_fread(imagesFile, &buffer4, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + imagesPath + " file");
        zip_fclose(imagesFile);
        if (document->Parse(buffer4, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + imagesPath + " file");
        parseRelationships(document);
        if ((options.flags >> 1) & 1) {
            saveImages();
        }
    }




}
