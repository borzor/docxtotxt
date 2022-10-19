//
// Created by borzor on 10/18/22.
//

#include <iostream>
#include "../headers/DocumentPropertiesParser.h"

using wstringConvert = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>;

void document::DocumentPropertiesParser::printData(const options_t &options) {
    tinyxml2::XMLDocument appXML, coreXML;
    auto input = options.input;
    int err = 0;
    struct zip_stat file_info{};
    zip_stat_init(&file_info);
    if (input == nullptr)
        throw std::runtime_error("Error: Cannot unzip file, error number: " + std::to_string(err));

    if (zip_stat(options.input, "docProps/app.xml", 0, &file_info) == -1)
        throw std::runtime_error("Error: Cannot get info about docProps/app.xml file");
    char buffer[file_info.size];
    auto appZip = zip_fopen(options.input, "docProps/app.xml", 0);
    if (zip_fread(appZip, &buffer, file_info.size) == -1)
        throw std::runtime_error("Error: Cannot read docProps/app.xml file");
    zip_fclose(appZip);
    if (appXML.Parse(buffer, file_info.size) != tinyxml2::XML_SUCCESS)
        throw std::runtime_error("Error: Cannot parse docProps/app.xml file");

    if (zip_stat(options.input, "docProps/core.xml", 0, &file_info) == -1)
        throw std::runtime_error("Error: Cannot get info about docProps/core.xml file");
    char buffer2[file_info.size];
    auto coreZip = zip_fopen(options.input, "docProps/core.xml", 0);
    if (zip_fread(coreZip, &buffer2, file_info.size) == -1)
        throw std::runtime_error("Error: Cannot read docProps/core.xml file");
    zip_fclose(coreZip);
    if (coreXML.Parse(buffer2, file_info.size) != tinyxml2::XML_SUCCESS)
        throw std::runtime_error("Error: Cannot parse docProps/app.xml file");
    switch (options.docType) {
        case pptx: {
            pptMetaData ppt;
            parsePptApp(appXML, ppt);
            parsePptCore(coreXML, ppt);
            printPptData(options, ppt);
            break;
        }
        case docx: {
            wordMetaData doc;
            parseWordApp(appXML, doc);
            parseWordCore(coreXML, doc);
            printWordData(options, doc);
            break;
        }
        case xlsx: {
            xlsxMetaData xlsx;
            parseXlsxApp(appXML, xlsx);
            parseXlsxCore(coreXML, xlsx);
            printXlsxData(options, xlsx);
            break;
        }
    }


}

void document::DocumentPropertiesParser::parseWordApp(tinyxml2::XMLDocument &appXML, wordMetaData &data) {
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

void document::DocumentPropertiesParser::parseWordCore(tinyxml2::XMLDocument &coreXML, wordMetaData &data) {
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

void document::DocumentPropertiesParser::printWordData(const options_t &options, wordMetaData &data) {
    *options.output << L"Document data:" << std::endl;
    *options.output << L"Revision №" << std::to_wstring(data.revision) << L", contains " << std::to_wstring(data.pages)
                    << " pages and " << std::to_wstring(data.words) << L" words" << std::endl;
    *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
    *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
    *options.output << L"Application: " << data.application << std::endl << std::endl;
    options.output->flush();
}


void document::DocumentPropertiesParser::parsePptApp(tinyxml2::XMLDocument &doc, pptMetaData &data) {
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

void document::DocumentPropertiesParser::parsePptCore(tinyxml2::XMLDocument &doc, pptMetaData &data) {
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

void document::DocumentPropertiesParser::printPptData(const options_t &options, pptMetaData &data) {
    *options.output << L"Document data:" << std::endl;
    *options.output << L"Revision №" << std::to_wstring(data.revision) << L", contains " << std::to_wstring(data.slides)
                    << " slides and " << std::to_wstring(data.words) << L" words" << std::endl;
    *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
    *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
    *options.output << L"Application: " << data.application << std::endl << std::endl;
    options.output->flush();
}

void document::DocumentPropertiesParser::parseXlsxApp(tinyxml2::XMLDocument &doc, xlsxMetaData &data) {
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

void document::DocumentPropertiesParser::parseXlsxCore(tinyxml2::XMLDocument &doc, xlsxMetaData &data) {
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

void document::DocumentPropertiesParser::printXlsxData(const options_t &options, xlsxMetaData &data) {
    *options.output << L"Document data:" << std::endl;
    *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
    *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
    *options.output << L"Application: " << data.application << std::endl << std::endl;
    options.output->flush();
}



