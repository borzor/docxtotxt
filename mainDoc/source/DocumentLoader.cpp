//
// Created by borzor on 11/30/22.
//

#include <numeric>
#include "../headers/DocumentLoader.h"
#include "../ParserCommons/CommonFunctions.h"

namespace docxtotxt {
    DocumentLoader::DocumentLoader(options_t &options, BufferWriter &writer) : options(options), writer(writer) {
        this->docInfo = {};
        this->pptInfo = {};
        this->xlsInfo = {};
        int err;
        options.input = zip_open(options.filePath, 0, &err);
        if (options.input == nullptr)throw runtime_error("Error: Cannot unzip file, error number: " + to_string(err));
    }

    void DocumentLoader::loadData() {
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
        if (((options.flags >> 4) & 1)) {
            openFileAndParse("docProps/app.xml", &DocumentLoader::parseAppFile);
            openFileAndParse("docProps/core.xml", &DocumentLoader::parseCoreFile);
        }
    }

    void DocumentLoader::openFileAndParse(const string &fileName, void (DocumentLoader::*f)(XMLDocument *)) {
        struct zip_stat file_info{};
        XMLDocument document;
        if (zip_stat(options.input, fileName.c_str(), 0, &file_info) == -1)
            throw runtime_error("Error: Cannot get info about " + fileName + " file");
        char *buff;
        buff = (char *) malloc(file_info.size);
        auto zip_file = zip_fopen(options.input, fileName.c_str(), 0);
        if (zip_fread(zip_file, buff, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + fileName + " file");
        zip_fclose(zip_file);
        if (document.Parse(buff, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + fileName + " file");
        ::free(buff);
        (this->*f)(&document);
    }

    void DocumentLoader::openFileAndParse(const string &fileName, relations_t &relations,
                                          void (DocumentLoader::*f)(XMLDocument *, relations_t &)) {
        struct zip_stat file_info{};
        XMLDocument document;
        if (zip_stat(options.input, fileName.c_str(), 0, &file_info) == -1)
            return; //its normal situation then no relations file
        char *buff;
        buff = (char *) malloc(file_info.size);
        auto zip_file = zip_fopen(options.input, fileName.c_str(), 0);
        if (zip_fread(zip_file, buff, file_info.size) == -1)
            throw runtime_error("Error: Cannot read " + fileName + " file");
        zip_fclose(zip_file);
        if (document.Parse(buff, file_info.size) != tinyxml2::XML_SUCCESS)
            throw runtime_error("Error: Cannot parse " + fileName + " file");
        ::free(buff);
        (this->*f)(&document, relations);
    }

    void DocumentLoader::parseAppFile(XMLDocument *doc) {
        auto *current_element = doc->RootElement()->FirstChildElement();
        writer.getMetadata()->application = L" ";
        switch (options.docType) {
            case pptx: {
                while (current_element != nullptr) {
                    if (!strcmp(current_element->Value(), "slides")) {
//                        writer.getMetadata()->slides = atoi(current_element->GetText());
                    } else if (!strcmp(current_element->Value(), "Words")) {
                        writer.getMetadata()->words = strtol(current_element->GetText(), nullptr, 10);
                    } else if (!strcmp(current_element->Value(), "Application")) {
                        auto text = current_element->GetText();
                        if (text == nullptr) {
                            //                            writer.getMetadata()->application = L" ";
                        } else
                            writer.getMetadata()->application = writer.convertString(text);
                    }
                    current_element = current_element->NextSiblingElement();
                }
                break;
            }
            case docx: {
                while (current_element != nullptr) {
                    if (!strcmp(current_element->Value(), "Pages")) {
                        writer.getMetadata()->pages = strtol(current_element->GetText(), nullptr, 10);
                    } else if (!strcmp(current_element->Value(), "Words")) {
                        writer.getMetadata()->words = strtol(current_element->GetText(), nullptr, 10);
                    } else if (!strcmp(current_element->Value(), "Application")) {
                        auto text = current_element->GetText();
                        if (text != nullptr)
                            writer.getMetadata()->application = writer.convertString(text);
                    }
                    current_element = current_element->NextSiblingElement();
                }
                break;
            }
            case xlsx: {
                while (current_element != nullptr) {
                    if (!strcmp(current_element->Value(), "Application")) {
                        auto text = current_element->GetText();
                        if (text != nullptr)
                            writer.getMetadata()->application = writer.convertString(text);
                    }
                    current_element = current_element->NextSiblingElement();
                }
                break;
            }
        }
    }

    void DocumentLoader::parseCoreFile(XMLDocument *doc) {
        wstringConvert convert;
        writer.getMetadata()->creator = L" ";
        writer.getMetadata()->lastModifiedBy = L" ";
        writer.getMetadata()->created = L" ";
        writer.getMetadata()->modified = L" ";
        auto *current_element = doc->RootElement()->FirstChildElement();
        while (current_element != nullptr) {
            if (!strcmp(current_element->Value(), "dc:creator")) {
                auto text = current_element->GetText();
                if (text != nullptr)
                    writer.getMetadata()->creator = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:lastModifiedBy")) {
                auto text = current_element->GetText();
                if (text != nullptr)
                    writer.getMetadata()->lastModifiedBy = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "cp:revision")) {
                writer.getMetadata()->revision = strtol(current_element->GetText(), nullptr, 10);
            } else if (!strcmp(current_element->Value(), "dcterms:created")) {
                auto text = current_element->GetText();
                if (text != nullptr)
                    writer.getMetadata()->created = convert.from_bytes(text);
            } else if (!strcmp(current_element->Value(), "dcterms:modified")) {
                auto text = current_element->GetText();
                if (text != nullptr)
                    writer.getMetadata()->modified = convert.from_bytes(text);
            }
            current_element = current_element->NextSiblingElement();
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

    void DocumentLoader::parseRelationShip(XMLDocument *doc, relations_t &relations) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        while (mainElement != nullptr) {
            if (!strcmp(mainElement->Value(), "Relationship")) {
                auto type = mainElement->Attribute("Type");
                string id = mainElement->Attribute("Id");
                string target = mainElement->Attribute("Target");
                if (ends_with(type, "image")) {
                    relations.imageRelationship.emplace(id, target.substr(target.find_last_of('/') + 1));
                } else if (ends_with(type, "hyperlink")) {
                    relations.hyperlinkRelationship.emplace(id, target);
                } else if (ends_with(type, "notesSlide") || (ends_with(type, "footnotes")) ||
                           (ends_with(type, "endnotes"))) {
                    relations.notes.emplace(id, target.substr(target.find_last_of('/') + 1));
                } else if (ends_with(type, "drawing")) {
                    relations.drawing.emplace(id, target.substr(target.find_last_of('/') + 1));
                }
            }
            mainElement = mainElement->NextSiblingElement();
        }
    }

    void DocumentLoader::parsePresentationSettings(XMLDocument *doc) {
        auto *element = doc->FirstChildElement()->FirstChildElement();
        while (element != nullptr) {
            if (!strcmp(element->Value(), "p:sldSz")) {
                auto width = strtol(element->Attribute("cx"), nullptr, 10);
                auto height = strtol(element->Attribute("cy"), nullptr, 10);
                pptInfo.settings.slideWidth = width;
                pptInfo.settings.slideHeight = height;
                pptInfo.settings.widthCoefficient = width / PRESENTATION_WIDTH;
                pptInfo.settings.heightCoefficient = height / PRESENTATION_HEIGHT;
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parsePresentationSlide(XMLDocument *doc) {
        auto *element = doc->RootElement()->FirstChildElement()->FirstChildElement();
        slideInfo slideInfo;
        while (element != nullptr) {
            if (!strcmp(element->Value(), "p:spTree")) {
                auto nodeElem = element->FirstChildElement();
                while (nodeElem != nullptr) {
                    if (!strcmp(nodeElem->Value(), "p:sp")) {
                        presentationText object;
                        parseSlideText(nodeElem, object);
                        slideInfo.objects.emplace_back(object);
                    } else if (!strcmp(nodeElem->Value(), "p:graphicFrame")) {
                        presentationTable object;
                        parseSlideTable(nodeElem, object);
                        slideInfo.tables.emplace_back(object);
                    } else if (!strcmp(nodeElem->Value(), "p:pic")) {
                        auto pic = extractPicture(nodeElem, "p");
                        slideInfo.pictures.emplace_back(pic);
                    }
                    nodeElem = nodeElem->NextSiblingElement();
                }
            }
            element = element->NextSiblingElement();
        }
        pptInfo.slides.emplace_back(slideInfo);
    }

    void DocumentLoader::parseSlideNote(XMLDocument *doc) {
        auto *element = doc->RootElement()->FirstChildElement()->FirstChildElement();
        slideInfo slideInfo;
        while (element != nullptr) {
            if (!strcmp(element->Value(), "p:spTree")) {
                auto nodeElem = element->FirstChildElement();
                while (nodeElem != nullptr) {
                    if (!strcmp(nodeElem->Value(), "p:sp")) {
                        auto txBody = nodeElem->FirstChildElement("p:txBody");
                        if (txBody != nullptr) {
                            auto tmp = extractTextBody(txBody);
                            pptInfo.notes.back().text.insert(pptInfo.notes.back().text.end(), tmp.begin(), tmp.end());
                        }
                    }
                    nodeElem = nodeElem->NextSiblingElement();
                }
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseSlideTable(XMLElement *element, presentationTable &object) {
        auto xfrm = element->FirstChildElement("p:xfrm");
        auto graphic = element->FirstChildElement("a:graphic");
        object.objectInfo = extractObjectInfo(xfrm);
        if (graphic != nullptr) {
            auto graphicData = graphic->FirstChildElement("a:graphicData");
            if (graphicData != nullptr) {
                string uri = graphicData->Attribute("uri");
                if (std::equal(uri.begin(), uri.end(), PPT_TABLE_URI)) {
                    auto tbl = graphicData->FirstChildElement("a:tbl");
                    if (tbl != nullptr) {
                        auto tblGrid = tbl->FirstChildElement("a:tblGrid");
                        auto tr = tbl->FirstChildElement("a:tr");
                        if (tblGrid != nullptr) {
                            auto gridCol = tblGrid->FirstChildElement("a:gridCol");
                            while (gridCol != nullptr) {
                                if (gridCol->Attribute("w") != nullptr) {
                                    object.gridColSize.emplace_back(
                                            strtol(gridCol->Attribute("w"), nullptr, 10) /
                                            pptInfo.settings.widthCoefficient);
                                }
                                gridCol = gridCol->NextSiblingElement("a:gridCol");
                            }
                        }
                        while (tr != nullptr) {
                            std::vector<std::vector<textBody>> tmpLine;
                            auto tc = tr->FirstChildElement("a:tc");
                            while (tc != nullptr) {
                                auto txBody = tc->FirstChildElement("a:txBody");
                                if (txBody != nullptr)
                                    tmpLine.emplace_back(extractTextBody(txBody));
                                tc = tc->NextSiblingElement("a:tc");
                            }
                            tr = tr->NextSiblingElement("a:tr");
                            object.table.emplace_back(tmpLine);
                        }
                    }
                }
            }
        }
    }

    void DocumentLoader::parseSlideText(XMLElement *element, presentationText &object) {
        auto spPr = element->FirstChildElement("p:spPr");
        auto txBody = element->FirstChildElement("p:txBody");
        if (spPr != nullptr) {
            auto xfrm = spPr->FirstChildElement("a:xfrm");
            object.objectInfo = extractObjectInfo(xfrm);
        }
        if (txBody != nullptr) {
            object.paragraph = extractTextBody(txBody);
        }
    }

    std::vector<textBody> DocumentLoader::extractTextBody(XMLElement *txBody) {
        auto par = txBody->FirstChildElement("a:p");
        std::vector<textBody> paragraph;
        while (par != nullptr) {
            textBody body;
            body.align = l;
            auto ppr = par->FirstChildElement("a:pPr");
            auto ar = par->FirstChildElement("a:r");
            while (ar != nullptr) {
                auto at = ar->FirstChildElement("a:t");
                if (at != nullptr && at->GetText() != nullptr) {
                    body.text += writer.convertString(at->GetText());
                }
                ar = ar->NextSiblingElement("a:r");
            }
            if (ppr != nullptr) {
                auto algn = ppr->Attribute("algn");
                if (algn != nullptr) {
                    if (!strcmp(algn, "r"))
                        body.align = r;
                    else if (!strcmp(algn, "ctr"))
                        body.align = ctr;
                }
                auto lvl = ppr->Attribute("lvl");
                auto buChar = ppr->FirstChildElement("a:buChar");
                if (buChar != nullptr) {
//                    auto char_ = buChar->Attribute("char");
//                    if (char_ != nullptr && !body.text.empty()) {
//                        body.text = std::wstring(1, ' ').append(L"•").append(body.text);
//                    }
                } else if (lvl != nullptr) {
                    auto sizeLvl = strtol(lvl, nullptr, 10);
                    if (!body.text.empty())
                        body.text = std::wstring(sizeLvl, ' ').append(L"•").append(body.text);
                }
            }
            if (!body.text.empty())
                paragraph.emplace_back(body);
            par = par->NextSiblingElement("a:p");
        }
        return paragraph;
    }

    void DocumentLoader::parseSharedStrings(XMLDocument *doc) {
        auto *element = doc->FirstChildElement()->FirstChildElement();
        while (element != nullptr) {
            if (!strcmp(element->Value(), "si")) {
                auto si = element->FirstChildElement();
                while (si != nullptr) {
                    if (!strcmp(si->Value(), "r")) { // pivotCacheRecords, convert all 'r' to one 't'
                        string tmpStr;
                        auto r = element->FirstChildElement();
                        while (r != nullptr) {
                            auto t = r->FirstChildElement("t");
                            if (t != nullptr && t->GetText() != nullptr) {
                                tmpStr += t->GetText();
                            }
                            r = r->NextSiblingElement();
                        }
                        xlsInfo.sharedStrings.emplace_back(writer.convertString(tmpStr));
                        break;
                    } else if (!strcmp(si->Value(), "t")) {
                        auto t = si->GetText();
                        if (t != nullptr)
                            xlsInfo.sharedStrings.emplace_back(writer.convertString(t));
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
                auto sheetName = element->Attribute("name");
                if (sheetName != nullptr)
                    resultSheet.sheetName = writer.convertString(sheetName);
                auto state = element->Attribute("state");
                if (state != nullptr)
                    resultSheet.state = writer.convertString(state);
                xlsInfo.worksheets.emplace_back(resultSheet);
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseDraw(XMLDocument *doc) {
        auto element = doc->RootElement()->FirstChildElement()->FirstChildElement("xdr:pic");
        auto pic = extractPicture(element, "xdr");
        xlsInfo.draws.back().pic = pic;
    }

    void DocumentLoader::parseWorksheet(XMLDocument *doc) {
        auto *element = doc->FirstChildElement()->FirstChildElement();
        sheet resultSheet;
        while (element != nullptr) {
            if (!strcmp(element->Value(), "sheetViews")) {
            } else if (!strcmp(element->Value(), "sheetPr") && element->Attribute("codeName") != nullptr) {
                resultSheet.sheetName = writer.convertString(element->Attribute("codeName"));
            } else if (!strcmp(element->Value(), "cols")) {
                auto col = element->FirstChildElement();
                while (col != nullptr) {
                    columnSettings tmpCol;
                    if (col->Attribute("min") != nullptr)
                        tmpCol.startInd = strtol(col->Attribute("min"), nullptr, 10) - 1;
                    if (col->Attribute("max") != nullptr)
                        tmpCol.endIndInd = strtol(col->Attribute("max"), nullptr, 10) - 1;
                    if (col->Attribute("width") != nullptr)
                        tmpCol.width = strtol(col->Attribute("width"), nullptr, 10);
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
                                    sheetCell tmpCell = {};
                                    if (row->Attribute("r") != nullptr) {
                                        string r = row->Attribute("r");
                                        size_t ii = 0, jj, colVal = 0;
                                        while (r[ii++] >= 'A') {}
                                        ii--;
                                        for (jj = 0; jj < ii; jj++)
                                            colVal = 26 * colVal + toupper(r[jj]) - 'A' + 1;
                                        tmpCell.cellNumber = colVal - 1;
                                    }
                                    if (row->Attribute("t") != nullptr)
                                        tmpCell.type = row->Attribute("t");
                                    if (row->FirstChildElement() != nullptr &&
                                        row->FirstChildElement()->GetText() != nullptr) {
                                        if (tmpCell.type == "s") {
                                            auto index = strtol(row->FirstChildElement()->GetText(), nullptr, 10);
                                            tmpCell.text = xlsInfo.sharedStrings[index];
                                        } else {
                                            tmpCell.text = writer.convertString(row->FirstChildElement()->GetText());
                                        }
                                    }
                                    tmpRow.emplace_back(tmpCell);
                                    row = row->NextSiblingElement();
                                }
                            }
                            resultSheet.sheetArray.emplace_back(tmpRow);
                        }
                    }
                    sheetData = sheetData->NextSiblingElement();
                }
                for (auto &sheet: xlsInfo.worksheets) {
                    if (sheet.sheetArray.empty()) {
                        sheet.sheetArray = resultSheet.sheetArray;
                        sheet.col = resultSheet.col;
                        break;
                    }
                }
            }
            element = element->NextSiblingElement();
        }
    }

    void DocumentLoader::parseStyles(XMLDocument *doc) {
        auto *mainElement = doc->RootElement()->FirstChildElement();
        if (options.docType == docx) {
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
                setIndentation(pPr->FirstChildElement("w:ind"), styleSettings.ind);
                setJustify(pPr->FirstChildElement("w:jc"), styleSettings.justify);
                setSpacing(pPr->FirstChildElement("w:spacing"), styleSettings.spacing);
                setTabulation(pPr->FirstChildElement("w:tabs"), styleSettings.tab);
            }
            if (def) {
                docInfo.styles.defaultStyles.paragraph = styleSettings;
            } else {
                docInfo.styles.paragraphStyles.emplace(styleId, styleSettings);
            }
        }
//        else if (!strcmp(type.c_str(), "table")) {
//            tableProperties styleSettings = {};
//            auto tblPr = element->FirstChildElement("w:tblPr");
//            if (tblPr != nullptr) {
//                setJustify(tblPr->FirstChildElement("w:tblInd"), styleSettings.justify);
//                docxtotxt::TableParser::setIndentation(tblPr->FirstChildElement("w:jc"), styleSettings.ind);
//                docxtotxt::TableParser::setFloatingSettings(tblPr->FirstChildElement("w:tblpPr"),
//                                                            styleSettings.floatTable);
//            }
//            if (def) {
//                docInfo.styles.defaultStyles.table = styleSettings;
//            } else {
//                docInfo.styles.tableStyles.emplace(styleId, styleSettings);
//            }
//        } else if (!strcmp(type.c_str(), "character")) {
//
//        } else if (!strcmp(type.c_str(), "numbering")) {
//
//        }
    }

    void DocumentLoader::setDefaultSettings(XMLElement *element) {
        auto *docDefault = element->FirstChildElement();
        docInfo.defaultSettings.paragraph.justify = justify_t::left;
        while (docDefault != nullptr) {
            if (!strcmp(docDefault->Value(), "w:rPrDefault")) {//default text properties

            } else if (!strcmp(docDefault->Value(), "w:pPrDefault")) {//default paragraph properties
                auto pPr = docDefault->FirstChildElement("w:pPr");
                if (pPr != nullptr) {
                    setIndentation(pPr->FirstChildElement("w:ind"),
                                   docInfo.defaultSettings.paragraph.ind);
                    setJustify(pPr->FirstChildElement("w:jc"),
                               docInfo.defaultSettings.paragraph.justify);
                }
            }
            docDefault = docDefault->NextSiblingElement();
        }
    }

    void DocumentLoader::loadDocxData() {
        string docStyles = DOC_STYLES_FILE;
        string mainFileNameProperty = DOC_MAIN_FILE;
        string footnoteProperty = DOC_FOOTNOTES_FILE;
        string endnoteProperty = DOC_ENDNOTES_FILE;
        openFileAndParse(DOC_IMAGE_FILE_PATH, docInfo.relations, &DocumentLoader::parseRelationShip);
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, docStyles)) {
                openFileAndParse(kv.first, &DocumentLoader::parseStyles);
            } else if (starts_with(kv.second, footnoteProperty) || starts_with(kv.second, endnoteProperty)) {
                openFileAndParse(kv.first, &DocumentLoader::parseDocNotes);
            }
        }
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, mainFileNameProperty)) {
                openFileAndParse(kv.first, &DocumentLoader::parseDocFile);
            }
        }
    }

    void DocumentLoader::loadXlsxData() {
        string sharedStrings = XLS_SHARED_STRINGS;
        string worksheet = XLS_WORKSHEET;
        string workbook = XLS_WORKBOOK;
        string drawing = XLS_SLIDE_NOTE;
        auto worksheetNum = 0;
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, sharedStrings)) {
                openFileAndParse(kv.first, &DocumentLoader::parseSharedStrings);
            } else if (starts_with(kv.second, worksheet)) {
                string path = kv.first + ".rels";
                path.insert(path.find_last_of('/', path.length()), "/_rels");
                openFileAndParse(kv.first, &DocumentLoader::parseWorksheet);
                openFileAndParse(path, xlsInfo.worksheets[worksheetNum].relations, &DocumentLoader::parseRelationShip);
                worksheetNum++;
            } else if (starts_with(kv.second, workbook)) {
                openFileAndParse(kv.first, &DocumentLoader::parseWorkbook);
            } else if (starts_with(kv.second, drawing)) {
                string path = kv.first + ".rels";
                path.insert(path.find_last_of('/', path.length()), "/_rels");
                draw draw;
                draw.name = kv.first.substr(kv.first.find_last_of('/') + 1);
                xlsInfo.draws.emplace_back(draw);
                openFileAndParse(kv.first, &DocumentLoader::parseDraw);
                openFileAndParse(path, xlsInfo.draws.back().relations, &DocumentLoader::parseRelationShip);
                //openFileAndParse(kv.first, &DocumentLoader::parseWorkbook);
            }
        }
    }

    void DocumentLoader::loadPptxData() {
        string slideContentType = PPT_SLIDE_CONTENT_TYPE;
        string presentationSettings = PPT_MAIN_FILE;
        string presentationNote = PPT_SLIDE_NOTE;
        for (const auto &kv: this->content_types) {
            if (starts_with(kv.second, slideContentType)) {
                string path = kv.first + ".rels";
                path.insert(path.find_last_of('/', path.length()), "/_rels");
                openFileAndParse(kv.first, &DocumentLoader::parsePresentationSlide);
                openFileAndParse(path, pptInfo.slides.back().relations, &DocumentLoader::parseRelationShip);
            } else if (starts_with(kv.second, presentationSettings)) {
                openFileAndParse(kv.first, &DocumentLoader::parsePresentationSettings);
            } else if (starts_with(kv.second, presentationNote)) {
                noteInfo_t noteInfo;
                noteInfo.noteName = kv.first.substr(kv.first.find_last_of('/') + 1);
                pptInfo.notes.emplace_back(noteInfo);
                openFileAndParse(kv.first, &DocumentLoader::parseSlideNote);
            }
        }
    }

    docInfo_t DocumentLoader::getDocxData() const {
        return docInfo;
    }

    xlsInfo_t DocumentLoader::getXlsxData() const {
        return this->xlsInfo;
    }

    pptInfo_t DocumentLoader::getPptxData() const {
        return pptInfo;
    }

    void DocumentLoader::parseSection(XMLElement *section) {
        XMLElement *sectionProperty = section->FirstChildElement("w:pgSz");
        if (sectionProperty != nullptr) {
            if (sectionProperty->Attribute("w:w") != nullptr)
                this->docInfo.docWidth = strtol(sectionProperty->Attribute("w:w"), nullptr, 10) / TWIP_TO_CHARACTER;
            if (sectionProperty->Attribute("w:h") != nullptr)
                this->docInfo.docHeight = strtol(sectionProperty->Attribute("w:h"), nullptr, 10) / TWIP_TO_CHARACTER;
        }
    }

    void DocumentLoader::parseDocFile(XMLDocument *doc) {
        XMLElement *section = doc->RootElement()->FirstChildElement()->FirstChildElement("w:sectPr");
        parseSection(section);
        auto mainElement = doc->RootElement()->FirstChildElement()->FirstChildElement();
        while (mainElement != nullptr) {
            paragraph par;
            par.type = paragraphType::par;
            if (!strcmp(mainElement->Value(), "w:p")) {
                parseParagraph(mainElement, par);
            } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                par.type = paragraphType::table;
                parseTable(mainElement, par, (options.flags >> 5) & 1);
            } else if (!strcmp(mainElement->Value(), "w:sdt")) {
                auto sdtContent = mainElement->FirstChildElement("w:sdtContent");
                if (sdtContent != nullptr) {
                    addSdtContent(sdtContent, docInfo.body);
                }
            }
            docInfo.body.emplace_back(par);
            mainElement = mainElement->NextSiblingElement();
        }
    }

    void DocumentLoader::parseParagraph(XMLElement *element, paragraph &par) {
        par.body.emplace_back();
        auto pPr = element->FirstChildElement("w:pPr");
        if (pPr == nullptr) {
            par.settings = docInfo.styles.defaultStyles.paragraph;
        } else {
            auto pStyle = pPr->FirstChildElement("w:pStyle");
            if (pStyle == nullptr) {
                par.settings = docInfo.styles.defaultStyles.paragraph;
            } else {
                par.settings = docInfo.styles.paragraphStyles[pStyle->Attribute("w:val")];
            }
            XMLElement *paragraphProperty = pPr->FirstChildElement();
            if (paragraphProperty->FirstAttribute())
                while (paragraphProperty != nullptr) {
                    if (!strcmp(paragraphProperty->Value(), "w:ind") &&
                        paragraphProperty->FirstAttribute() != nullptr) {
                        setIndentation(paragraphProperty, par.settings.ind);
                    } else if (!strcmp(paragraphProperty->Value(), "w:jc") &&
                               paragraphProperty->FirstAttribute() != nullptr) {
                        setJustify(paragraphProperty, par.settings.justify);
                    } else if (!strcmp(paragraphProperty->Value(), "w:numPr")) {
                        XMLElement *enumProperty = paragraphProperty->FirstChildElement();
                        if (!strcmp(enumProperty->Value(), "w:ilvl")) {
                            if (enumProperty->FirstAttribute() != nullptr) {
                                size_t indentationSize = enumProperty->FirstAttribute()->IntValue();
                                par.body.back().append(
                                        writer.convertString(std::string(indentationSize, ' ') + (" · ")));
                            }
                        }
                        setJustify(enumProperty, par.settings.justify);
                    } else if (!strcmp(paragraphProperty->Value(), "w:outlineLvl") &&
                               paragraphProperty->FirstAttribute() != nullptr) {
                        par.settings.outline = true;
                    } else if (!strcmp(paragraphProperty->Value(), "w:spacing") &&
                               paragraphProperty->FirstAttribute() != nullptr) {
                        setSpacing(paragraphProperty, par.settings.spacing);
                    } else if (!strcmp(paragraphProperty->Value(), "w:tabs") &&
                               paragraphProperty->FirstChildElement() != nullptr) {
                        setTabulation(paragraphProperty, par.settings.tab);
                    }
                    paragraphProperty = paragraphProperty->NextSiblingElement();
                }
        }
        XMLElement *property = element->FirstChildElement();
        while (property != nullptr) {
            if (!strcmp(property->Value(), "w:r")) {
                parseTextProperties(property, par);
            } else if (!strcmp(property->Value(), "w:hyperlink")) {
                XMLElement *linkProp = property->FirstChildElement();
                while (linkProp != nullptr) {
                    if (!strcmp(linkProp->Value(), "w:r")) {
                        parseTextProperties(linkProp, par);
                    }
                    linkProp = linkProp->NextSiblingElement();
                }
                if ((options.flags >> 2) & 1) {
                    if (property->Attribute("r:id") != nullptr) {
                        auto id = property->Attribute("r:id");
                        auto number = distance(docInfo.relations.hyperlinkRelationship.begin(),
                                               docInfo.relations.hyperlinkRelationship.find(id));//very doubtful
                        auto result = string("{h").append(to_string(number).append("}"));
                        par.body.back().append(writer.convertString(result));
                    }
                }
            } else {
                //throw runtime_error(string("Unexpected child element: ") + string(property->Value()));
            }
            property = property->NextSiblingElement();
        }
        free(property);
    }

    void DocumentLoader::parseTextProperties(XMLElement *properties, paragraph &par) {
        XMLElement *textProperty = properties->FirstChildElement();
        while (textProperty != nullptr) {
            if (!strcmp(textProperty->Value(), "w:drawing")) {
                auto current_element = textProperty->FirstChildElement()->FirstChildElement();
                while (current_element != nullptr) {
                    if (!strcmp(current_element->Value(), "a:graphic")) {
                        if (current_element->FirstChildElement("a:graphicData") != nullptr) {
                            if (current_element->FirstChildElement("a:graphicData")->FirstChildElement(
                                    "pic:pic") != nullptr) {
                                auto picElement = current_element->FirstChildElement(
                                        "a:graphicData")->FirstChildElement(
                                        "pic:pic");
                                auto pic = extractPicture(picElement, "pic");
                                par.type = paragraphType::image;
                                auto imageName = docInfo.relations.imageRelationship[pic.rId];
                                auto width = pic.objectInfo.objectSizeX / 76200;
                                auto height = pic.objectInfo.objectSizeY / 76200;
                                auto mediumLine = width - 2;
                                auto leftBorder = (docInfo.docWidth - width) / 2;
                                wstring path = L"Media file.";
                                par.body.emplace_back(path);
                                if ((width < 3 || height < 3)) {
                                    wstring imageInfo = wstring(L" Image is too small to display");
                                    par.body.back().append(imageInfo);
                                }
                                if ((options.flags >> 1) & 1) {
                                    wstring imageInfo = wstring(L"Saved to path: ") +
                                                        writer.convertString(
                                                                this->options.pathToDraws + '/' + imageName);
                                    par.body.emplace_back(imageInfo);
                                }
                                if (width > 3 || height > 3) {
                                    par.body.emplace_back(
                                            std::wstring(1, L'+').append(mediumLine, L'—').append(1, L'+'));
                                    for (int i = 1; i < height - 1; i++) {
                                        par.body.emplace_back(L"|" + std::wstring(mediumLine, L' ') + L"|");
                                    }
                                    par.body.emplace_back(
                                            std::wstring(1, L'+').append(mediumLine, L'—').append(1, L'+'));

                                }
                                par.body.emplace_back();
                            }
                        }

                    }
                    current_element = current_element->NextSiblingElement();
                }
            } else if (!strcmp(textProperty->Value(), "w:t")) {
                auto text = textProperty->GetText();
                if (text == nullptr)
                    par.body.back().append(L" ");
                else
                    par.body.back().append(writer.convertString(text));
            } else if (!strcmp(textProperty->Value(), "w:tab")) {
                if (par.settings.tab.empty()) {
                    par.body.back().append(L"   ");
                } else {
                    auto tab = par.settings.tab.front();
                    auto pos = tab.pos;
                    char symbol = ' ';
                    if (tab.character == dot)
                        symbol = '.';
                    auto currentSize = par.body.back().length();
                    auto ind = par.settings.ind.left;
                    if (ind != 0)
                        currentSize += ind;
                    if (pos > currentSize)
                        par.body.back().append(pos - currentSize, symbol);
                    par.settings.tab.erase(par.settings.tab.begin());
                }
            } else if(!strcmp(textProperty->Value(), "w:footnoteReference")) {
                auto id = textProperty->Attribute("w:id");
                if(id!= nullptr)
                    par.body.back().append(L"{footnote" + writer.convertString(id) + L"}");
            } else if(!strcmp(textProperty->Value(), "w:endnoteReference")) {
                auto id = textProperty->Attribute("w:id");
                if(id!= nullptr)
                    par.body.back().append(L"{endnote" + writer.convertString(id) + L"}");
            }
            textProperty = textProperty->NextSiblingElement();
        }
    }

    void DocumentLoader::parseTable(XMLElement *tableElem, ::docxtotxt::paragraph &par, const bool raw) {
        XMLElement *element = tableElem->FirstChildElement();
        tableSetting settings;
        size_t line = 0;
        size_t currentColumn = 0;
        while (element != nullptr) {
//            if (!strcmp(element->Value(), "w:tblPr")) {// table Properties
//                auto tblStyle = element->FirstChildElement("w:tblStyle");
//                if (tblStyle == nullptr) {
//                    settings.settings = docInfo.styles.defaultStyles.table;
//                } else {
//                    settings.settings = docInfo.styles.tableStyles[tblStyle->Attribute("w:val")];
//                }
//                XMLElement *tableProperty = element->FirstChildElement();
//                while (tableProperty != nullptr) {
//                    if (!strcmp(tableProperty->Value(), "w:tblInd")) {
//                        if (tableProperty->FirstAttribute() != nullptr) {
//                            auto tblInd = tableProperty->Attribute("w:w");
//                            settings.settings.ind = atoi(tblInd);
//                        }
//                    } else if (!strcmp(tableProperty->Value(), "w:tblpPr")) {
//
//                    }
//                    tableProperty = tableProperty->NextSiblingElement();
//                }
//            } else
            if (!strcmp(element->Value(), "w:tblGrid")) {
                XMLElement *gridCol = element->FirstChildElement();
                while (gridCol != nullptr) {
                    settings.tblGrids.emplace_back(std::stoi(gridCol->Attribute("w:w")));
                    gridCol = gridCol->NextSiblingElement();
                }
                settings.columnAmount = settings.tblGrids.size();
            } else if (!strcmp(element->Value(), "w:tr")) {//table row
                XMLElement *rowElement = element->FirstChildElement();
                settings.grid.emplace_back();
                while (rowElement != nullptr) {
                    if (!strcmp(rowElement->Value(), "w:tc")) {//Specifies a table cell
                        XMLElement *tcElement = rowElement->FirstChildElement();
                        while (tcElement != nullptr) {
                            if (!strcmp(tcElement->Value(), "w:p")) {//
                                paragraph tmpPar;
                                parseParagraph(tcElement, tmpPar);
                                settings.grid[line].back().text += tmpPar.body.back();
                                //getTableCellParagraph(tcElement, paragraphParser); //TODO
                            } else if (!strcmp(tcElement->Value(), "w:tbl")) {//
                                // for the table in table...
                            } else if (!strcmp(tcElement->Value(), "w:tcPr")) {//Specifies a table cell
                                XMLElement *gridCol = tcElement->FirstChildElement();
                                size_t gridSpan = 1;
                                cell tmpCell;
                                while (gridCol != nullptr) {
                                    if (!strcmp(gridCol->Value(), "w:tcW")) {
                                        auto size = gridCol->Attribute("w:w");
                                        if (size != nullptr) {
                                            tmpCell.width = strtol(size, nullptr, 10) / TWIP_TO_CHARACTER;
                                        }
                                    } else if (!strcmp(gridCol->Value(), "w:gridSpan")) {
                                        gridSpan = strtol(gridCol->Attribute("w:val"), nullptr, 10);
                                    }
                                    gridCol = gridCol->NextSiblingElement();
                                }
                                tmpCell.gridSpan = gridSpan;
                                settings.grid[line].push_back(tmpCell);
                                free(gridCol);
                            }
                            tcElement = tcElement->NextSiblingElement();
                        }
                        currentColumn += settings.grid[line].back().gridSpan;
                    }
                    rowElement = rowElement->NextSiblingElement();
                }
                currentColumn = 0;
                line += 1;
            }
            element = element->NextSiblingElement();
        }

        line = 0;
        size_t tableWidth = std::accumulate(settings.grid[line].begin(), settings.grid[line].end(), 0,
                                            [](size_t tableWidth, const cell &tmpCell) {
                                                return tableWidth + tmpCell.width;
                                            });
        size_t minColumn = settings.columnAmount;
        for (auto &i: settings.grid) {
            minColumn = min(minColumn, i.size());
        }
        auto tableSize = tableWidth;
        tableSize += (minColumn + 1); //column size + amount of columns
        auto mediumLine = tableSize - 2;
        if (raw) {
            std::vector<size_t> columnSize(settings.columnAmount, 0);
            for (auto &row: settings.grid) {
                for (int cell = 0; cell < row.size(); cell++) {
                    columnSize[cell] = std::max(columnSize[cell], row[cell].text.size());
                }
            }
            for (auto &row: settings.grid) {
                par.body.emplace_back();
                for (int cell = 0; cell < row.size(); cell++) {
                    par.body.back().append(row[cell].text).append(
                            std::wstring(columnSize[cell] - row[cell].text.size() + 1, L' '));
                }
            }
        } else {
            par.body.emplace_back(std::wstring(1, L'+').append(mediumLine, L'—').append(1, L'+'));
            while (line < settings.grid.size()) {
                bool lineDone = true;
                currentColumn = 0;
                par.body.emplace_back(L"|");
                while (currentColumn < settings.grid[line].size()) {
                    auto charInCell = settings.grid[line][currentColumn].width;
                    if (!settings.grid[line][currentColumn].text.empty()) {
                        auto text = settings.grid[line][currentColumn].text;
                        std::wstring resultText;
                        auto indexLastElement = text.find_last_of(L' ', charInCell);
                        if (charInCell < text.size()) {
                            lineDone = false;
                            if (indexLastElement == string::npos) {
                                resultText = text.substr(0, charInCell);
                            } else {
                                resultText = text.substr(0, indexLastElement);
                            }
                            par.body.back().append(resultText);
                            par.body.back().append(indexLastElement == string::npos ? 0 : charInCell - indexLastElement,
                                                   L' ');
                        } else {
                            resultText = text;
                            par.body.back().append(resultText);
                            par.body.back().append(charInCell - text.size(), L' ');
                        }
                        settings.grid[line][currentColumn].text.erase(0, resultText.size() + 1);
                        par.body.back().append(1, L'|');
                    } else {
                        par.body.back().append(charInCell == 1 ? 0 : charInCell, L' ').append(1, L'|');
                    }
                    currentColumn++;
                }
                if (lineDone) {
                    line++;
                    par.body.emplace_back(std::wstring(1, L'+').append(mediumLine, L'—').append(1, L'+'));
                }
            }
        }
    }

    void DocumentLoader::addSdtContent(XMLElement *sdtContent, vector<paragraph> &body) {
        auto childElem = sdtContent->FirstChildElement();
        while (childElem != nullptr) {
            if (!strcmp(childElem->Value(), "w:p")) {
                paragraph par;
                par.type = paragraphType::par;
                parseParagraph(childElem, par);
                body.emplace_back(par);
            } else if (!strcmp(childElem->Value(), "w:tbl")) {
                paragraph par;
                par.type = paragraphType::table;
                parseTable(childElem, par, true);
                body.emplace_back(par);
            } else if (!strcmp(childElem->Value(), "w:sdt")) {
                auto sdt = childElem->FirstChildElement("w:sdtContent");
                if (sdt != nullptr) {
                    addSdtContent(sdt, body);
                }
            }
            childElem = childElem->NextSiblingElement();
        }
    }

    void DocumentLoader::parseDocNotes(XMLDocument *doc) {
        note object;
        object.text = L"";
        XMLElement *childElem = doc->FirstChildElement();
        if (childElem->FirstChildElement("w:footnote") != nullptr) {
            childElem = childElem->FirstChildElement("w:footnote");
            object.type = footnote;
        } else if (childElem->FirstChildElement("w:endnote") != nullptr) {
            childElem = childElem->FirstChildElement("w:endnote");
            object.type = endnote;
        } else
            return;
        while (childElem != nullptr) {
            auto id = childElem->Attribute("w:id");
            if (id != nullptr)
                object.id = strtol(id, nullptr, 10);
            auto p = childElem->FirstChildElement("w:p");
            while (p != nullptr) {
                auto r = p->FirstChildElement("w:r");
                while (r != nullptr) {
                    auto t = r->FirstChildElement("w:t");
                    if (t != nullptr && t->GetText() != nullptr) {
                        object.text += writer.convertString(t->GetText());
                    }
                    r = r->NextSiblingElement();
                }
                p = p->NextSiblingElement();
                if (!object.text.empty())
                    docInfo.notes.emplace_back(object);
            }
            childElem = childElem->NextSiblingElement();
        }
    }


}








