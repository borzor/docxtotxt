//
// Created by borzor on 11/30/22.
//

#include "../headers/DocumentLoader.h"
#include "../headers/ParagraphParser.h"
#include "../headers/TableParser.h"

namespace Loader {
    DocumentLoader::DocumentLoader(options_t &options) : options(options) {
        int err;
        options.input = zip_open(options.filePath, 0, &err);
        this->docInfo.resultBuffer.pointer = 0;
        if (options.input == nullptr)
            throw runtime_error("Error: Cannot unzip file, error number: " + to_string(err));
    }

    void DocumentLoader::loadData() {
        string mainFileNameProperty, defaultMainPath, mainFileStyleProperty, defaultStylePath;
        openFileAndParse(CONTENT_TYPE_NAME, &DocumentLoader::parseContentTypes);
        switch (options.docType) {
            case pptx: {
                loadPptxData();
                break;
            }
            case docx: {
                loadDocxData();
                break;
            }
            case xlsx: {
                loadXlsxData();
                break;
            }
        }
        if (options.printDocProps) {
            openFileAndParse("docProps/app.xml", &DocumentLoader::parseAppFile);
            openFileAndParse("docProps/core.xml", &DocumentLoader::parseCoreFile);
        }
    }

    void DocumentLoader::openFileAndParse(const string &fileName, void (DocumentLoader::*f)(XMLDocument *)) {
        struct zip_stat file_info{};
        XMLDocument document;
        if (zip_stat(options.input, fileName.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + fileName + " file");
        char buffer4[file_info.size];
        auto zip_file = zip_fopen(options.input, fileName.c_str(), 0);
        if (zip_fread(zip_file, &buffer4, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + fileName + " file");
        zip_fclose(zip_file);
        if (document.Parse(buffer4, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + fileName + " file");
        (this->*f)(&document);
        //zip_fclose(zip_file);
    }


    void DocumentLoader::parseAppFile(XMLDocument *doc) {
        switch (options.docType) {
            case pptx: {
                parsePptApp(*doc, pptInfo.fileMetaData);
                break;
            }
            case docx: {
                parseWordApp(*doc, docInfo.fileMetaData);
                break;
            }
            case xlsx: {
                parseXlsxApp(*doc, xlsInfo.fileMetaData);
                break;
            }
        }
    }

    void DocumentLoader::parseCoreFile(XMLDocument *doc) {
        switch (options.docType) {
            case pptx: {
                parsePptCore(*doc, pptInfo.fileMetaData);
                break;
            }
            case docx: {
                parseWordCore(*doc, docInfo.fileMetaData);
                break;
            }
            case xlsx: {
                parseXlsxCore(*doc, xlsInfo.fileMetaData);
                break;
            }
        }
    }

    void DocumentLoader::parseContentTypes(XMLDocument *xmlDocument) {
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

    void DocumentLoader::parseRelationships(XMLDocument *doc) {
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

    void DocumentLoader::parseSharedStrings(XMLDocument *doc) {
        auto *element = doc->FirstChildElement()->FirstChildElement();
        while (element != nullptr) {
            if (!strcmp(element->Value(), "si")) {
                auto si = element->FirstChildElement();
                while (si != nullptr) {
                    if (!strcmp(si->Value(), "r")) { // pivotCacheRecords
                        auto r = element->FirstChildElement();
                        break;
                    } else if (!strcmp(si->Value(), "t")) {
                        auto t = si->GetText();
                        if (t != nullptr)
                            xlsInfo.sharedStrings.emplace_back(convertor.from_bytes(t));
                        else
                            xlsInfo.sharedStrings.emplace_back(L" ");
                    }
                    si = si->NextSiblingElement();
                }
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseWorkbook(XMLDocument *doc) {
        auto element = doc->RootElement()->FirstChildElement("sheets")->FirstChildElement();
        while (element != nullptr) {
            sheet resultSheet;
            resultSheet.sheetName = L"";
            resultSheet.state = L"";
            if (!strcmp(element->Value(), "sheet")) {
                auto sheetName =  element->Attribute("name");
                if(sheetName != nullptr)
                    resultSheet.sheetName = convertor.from_bytes(sheetName);
                auto state = element->Attribute("state");
                if(state != nullptr)
                    resultSheet.state = convertor.from_bytes(state);
                if (this->xlsInfo.worksheets.count(resultSheet.sheetName)) {
                    this->xlsInfo.worksheets[resultSheet.sheetName].sheetName = resultSheet.sheetName;
                    this->xlsInfo.worksheets[resultSheet.sheetName].state = resultSheet.state;
                } else {
                    xlsInfo.worksheets.insert(std::pair<std::wstring, sheet>(resultSheet.sheetName, resultSheet));
                }
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseWorksheet(XMLDocument *doc) {
        auto *element = doc->FirstChildElement()->FirstChildElement();
        sheet resultSheet;
        while (element != nullptr) {
            if (!strcmp(element->Value(), "sheetViews")) {
            } else if (!strcmp(element->Value(), "sheetPr")) {
                resultSheet.sheetName = convertor.from_bytes(element->Attribute("codeName"));
            } else if (!strcmp(element->Value(), "cols")) {
                std::vector<columnSettings> tmpColSetting;
                auto col = element->FirstChildElement();
                while (col != nullptr) {
                    columnSettings tmpCol;
                    if (col->Attribute("min") != nullptr)
                        tmpCol.startInd = atoi(col->Attribute("min")) - 1;
                    if (col->Attribute("max") != nullptr)
                        tmpCol.endIndInd = atoi(col->Attribute("max")) - 1;
                    if (col->Attribute("width") != nullptr)
                        tmpCol.width = atoi(col->Attribute("width"));
                    resultSheet.col.emplace_back(tmpCol);
                    col = col->NextSiblingElement();
                }
            } else if (!strcmp(element->Value(), "sheetData")) {
                auto sheetData = element->FirstChildElement();
                while (sheetData != nullptr) {
                    if (!strcmp(sheetData->Value(), "row")) {
                        std::vector<sheetCell> tmpRow;
                        auto row = sheetData->FirstChildElement();
                        while (row != nullptr) {
                            if (!strcmp(row->Value(), "c")) {
                                while (row != nullptr) {
                                    sheetCell tmpCell;
                                    if (row->Attribute("r") != nullptr) {
                                        string r = row->Attribute("r");
                                        size_t ii = 0, jj, colVal = 0;
                                        while (r[ii++] >= 'A') {};
                                        ii--;
                                        for (jj = 0; jj < ii; jj++)
                                            colVal = 26 * colVal + toupper(r[jj]) - 'A' + 1;
                                        tmpCell.cellNumber = colVal - 1;
                                    }
                                    if (row->Attribute("t") != nullptr)
                                        tmpCell.type = row->Attribute("t");//TODO add switch types ?
                                    if (row->FirstChildElement() != nullptr &&
                                        row->FirstChildElement()->GetText() != nullptr)
                                        if (tmpCell.type == "s") {
                                            tmpCell.text = xlsInfo.sharedStrings[
                                                    atoi(row->FirstChildElement()->GetText()) - 1];
                                        } else {
                                            tmpCell.text = convertor.from_bytes(row->FirstChildElement()->GetText());
                                        }
                                    tmpRow.emplace_back(tmpCell);
                                    row = row->NextSiblingElement();
                                }
                            }
                            resultSheet.sheetArray.emplace_back(tmpRow);
                            sheetData = sheetData->NextSiblingElement();
                        }
                    }
                }
                if (this->xlsInfo.worksheets.count(resultSheet.sheetName)) {
                    this->xlsInfo.worksheets[resultSheet.sheetName].sheetArray = resultSheet.sheetArray;
                    this->xlsInfo.worksheets[resultSheet.sheetName].col = resultSheet.col;
                } else {
                    xlsInfo.worksheets.insert(std::pair<std::wstring, sheet>(resultSheet.sheetName, resultSheet));
                }
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseStyles(XMLDocument *doc) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        switch (options.docType) {
            case pptx:
                break;
            case docx: {
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

    void DocumentLoader::addStyle(XMLElement *element) {
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

    void DocumentLoader::setDefaultSettings(XMLElement *element) {
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

    void DocumentLoader::parseWordApp(tinyxml2::XMLDocument &appXML, fileMetaData_t &data) {
        auto *current_element = appXML.RootElement()->FirstChildElement();
        wstringConvert convert;
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "Pages")) {
                data.pages = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "Words")) {
                data.words = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "Application")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.application = L" ";
                else
                    data.application = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseWordCore(tinyxml2::XMLDocument &coreXML, fileMetaData_t &data) {
        wstringConvert convert;
        auto *current_element = coreXML.RootElement()->FirstChildElement();
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "dc:creator")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.creator = L" ";
                else
                    data.creator = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:lastModifiedBy")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.lastModifiedBy = L" ";
                else
                    data.lastModifiedBy = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:revision")) {
                data.revision = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "dcterms:created")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.created = L" ";
                else
                    data.created = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "dcterms:modified")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.modified = L" ";
                else
                    data.modified = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::parsePptApp(tinyxml2::XMLDocument &doc, fileMetaData_t &data) {
        auto *current_element = doc.RootElement()->FirstChildElement();
        wstringConvert convert;
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "slides")) {
                data.slides = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "Words")) {
                data.words = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "Application")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.application = L" ";
                else
                    data.application = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::parsePptCore(tinyxml2::XMLDocument &doc, fileMetaData_t &data) {
        wstringConvert convert;
        auto *current_element = doc.RootElement()->FirstChildElement();
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "dc:creator")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.creator = L" ";
                else
                    data.creator = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:lastModifiedBy")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.lastModifiedBy = L" ";
                else
                    data.lastModifiedBy = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:revision")) {
                data.revision = atoi(current_element->GetText());
            } else if (!strcmp(current_element->Value(), "dcterms:created")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.created = L" ";
                else
                    data.created = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "dcterms:modified")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.modified = L" ";
                else
                    data.modified = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseXlsxApp(tinyxml2::XMLDocument &doc, fileMetaData_t &data) {
        auto *current_element = doc.RootElement()->FirstChildElement();
        wstringConvert convert;
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "Application")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.application = L" ";
                else
                    data.application = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseXlsxCore(tinyxml2::XMLDocument &doc, fileMetaData_t &data) {
        wstringConvert convert;
        auto *current_element = doc.RootElement()->FirstChildElement();
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "dc:creator")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.creator = L" ";
                else
                    data.creator = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:lastModifiedBy")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.lastModifiedBy = L" ";
                else
                    data.lastModifiedBy = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "dcterms:created")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.created = L" ";
                else
                    data.created = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "dcterms:modified")) {
                auto text = current_element->GetText();
                if (text == nullptr)
                    data.modified = L" ";
                else
                    data.modified = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DocumentLoader::loadDocxData() {
        string docStyles = DOC_STYLES_FILE;
        string mainFileNameProperty = DOC_MAIN_FILE;
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, docStyles)) {
                openFileAndParse(kv.first, &DocumentLoader::parseStyles);
            } else if (starts_with(kv.second, mainFileNameProperty)) {
                string fileName = kv.first;
                struct zip_stat file_info{};
                if (zip_stat(options.input, fileName.c_str(), 0, &file_info) == -1)
                    throw runtime_error("Error: Cannot get info about " + fileName + " file");
                char buffer4[file_info.size];
                auto zip_file = zip_fopen(options.input, fileName.c_str(), 0);
                if (zip_fread(zip_file, &buffer4, file_info.size) == -1)
                    throw runtime_error("Error: Cannot read " + fileName + " file");
                zip_fclose(zip_file);
                if (mainDoc.Parse(buffer4, file_info.size) != tinyxml2::XML_SUCCESS)
                    throw runtime_error("Error: Cannot parse " + fileName + " file");
            }
        }
        openFileAndParse(DOC_IMAGE_FILE_PATH, &DocumentLoader::parseRelationships);
    }

    void DocumentLoader::loadXlsxData() {
        string sharedStrings = XLS_SHARED_STRINGS;
        string worksheet = XLS_WORKSHEET;
        string workbook = XLS_WORKBOOK;
        xlsInfo.sharedStrings.emplace_back(L"firstEmptyElement");
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, sharedStrings)) {
                openFileAndParse(kv.first, &DocumentLoader::parseSharedStrings);
            } else if (starts_with(kv.second, worksheet)) {
                openFileAndParse(kv.first, &DocumentLoader::parseWorksheet);
            } else if (starts_with(kv.second, workbook)) {
                openFileAndParse(kv.first, &DocumentLoader::parseWorkbook);
            }
        }
    }

    void DocumentLoader::loadPptxData() {
        struct zip_stat file_info{};
        string slideContentType = PPT_SLIDE_CONTENT_TYPE;
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, slideContentType)) {
                //TODO openFileAndParse(kv.first, &MainDocParser::parseWorksheet);
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

    docInfo_t DocumentLoader::getDocxData() {
        return docInfo;
    }

    XMLDocument *DocumentLoader::getMainDocument() {
        return &mainDoc;
    }

    xlsInfo_t DocumentLoader::getXlsxData() const {
        return this->xlsInfo;
    }

    pptInfo_t DocumentLoader::getPptxData() const {
        pptInfo_t docInfo1;
        return docInfo1;
    }


}