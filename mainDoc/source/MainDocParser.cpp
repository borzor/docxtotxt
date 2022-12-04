//
// Created by boris on 2/9/22.
//

#include <string>
#include <locale>
#include <regex>

#include "../headers/MainDocParser.h"
#include "../headers/TableParser.h"
#include "../headers/XlsParser.h"
#include "../headers/DocumentLoader.h"

namespace prsr {
    void MainDocParser::parseFile() {
        Loader::DocumentLoader documentLoader(options);
        documentLoader.loadData();
        switch (options.docType) {
            case pptx: {
                auto pptInfo = documentLoader.getPptxData();
                if (options.printDocProps) {
                    printMetaInfoData(pptInfo.fileMetaData);
                }
                break;
            }
            case docx: {
                auto docInfo = documentLoader.getDocxData();
                if (options.printDocProps) {
                    printMetaInfoData(docInfo.fileMetaData);
                }
                auto a = documentLoader.getMainDocument();
                auto mainElement = a->RootElement()->FirstChildElement()->FirstChildElement();
                XMLElement *section = a->RootElement()->FirstChildElement()->FirstChildElement("w:sectPr");
                auto sectionParser = section::SectionParser(docInfo);
                sectionParser.parseSection(section);
                paragraph::ParagraphParser paragraphParser(docInfo, options);
                auto tableParser = table::TableParser(docInfo, options);
                addLine(docInfo.resultBuffer);
                while (mainElement != nullptr) {
                    try {
                        if (!strcmp(mainElement->Value(), "w:p")) {
                            paragraphParser.parseParagraph(mainElement);
                            paragraphParser.writeResult();
                            paragraphParser.flush();
                        } else if (!strcmp(mainElement->Value(), "w:tbl")) {
                            tableParser.parseTable(mainElement);
                            tableParser.insertTable();
                            tableParser.flush();
                        } else {
                            //throw runtime_error(string("Unexpected main element: ") + string(mainElement->Value()));
                        }
                        mainElement = mainElement->NextSiblingElement();
                    } catch (exception &e) {
                        std::cout << "Allocation failed: " << e.what() << '\n';
                    }
                }
                printResult(docInfo.resultBuffer);
                if ((options.flags >> 1) & 1) {
                    saveImages(docInfo.imageRelationship);
                }
                if ((options.flags >> 2) & 1) {
                    insertHyperlinks(docInfo.hyperlinkRelationship);
                }
                free(mainElement);
                break;
            }
            case xlsx: {
                auto xlsInfo = documentLoader.getXlsxData();
                if (options.printDocProps) {
                    printMetaInfoData(xlsInfo.fileMetaData);
                }
                xls::XlsParser xlsParser(xlsInfo, options);
                xlsParser.parseSheets();
                printResult(xlsInfo.resultBuffer);
                break;
            }
        }
        zip_close(options.input); //TODO remake it

    }

    MainDocParser::MainDocParser(options_t &options) :
            options(options) {

    }

    void MainDocParser::insertHyperlinks(std::map<std::string, std::string> hyperlinkRelationship) {
        for (const auto &kv: hyperlinkRelationship) {
            auto number = distance(hyperlinkRelationship.begin(),
                                   hyperlinkRelationship.find(kv.first));//very doubtful
            auto result = wstring(L"{h").append(
                    to_wstring(number).append(L"} -  ").append(convertor.from_bytes(kv.second)));
            *options.output << result << '\n';
            options.output->flush();
        }
    }

    void MainDocParser::saveImages(const std::map<std::string, std::string> &imageRelationship) const {
        struct zip_stat file_info{};
        zip_stat_init(&file_info);
        for (const auto &kv: imageRelationship) {
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


    void MainDocParser::printResult(buffer &resultBuffer) const {
        for (const auto &kv: resultBuffer.buffer) {
            *options.output << kv << std::endl;
        }
    }

    void MainDocParser::printMetaInfoData(const fileMetaData_t &data) const {
        switch (options.docType) {
            case pptx: {
                *options.output << L"Document data:" << std::endl;
                *options.output << L"Revision №" << std::to_wstring(data.revision) << L", contains "
                                << std::to_wstring(data.slides)
                                << " slides and " << std::to_wstring(data.words) << L" words" << std::endl;
                *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
                *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
                *options.output << L"Application: " << data.application << std::endl << std::endl;
                options.output->flush();
                break;
            }
            case docx: {
                *options.output << L"Document data:" << std::endl;
                *options.output << L"Revision №" << std::to_wstring(data.revision) << L", contains "
                                << std::to_wstring(data.pages)
                                << " pages and " << std::to_wstring(data.words) << L" words" << std::endl;
                *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
                *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
                *options.output << L"Application: " << data.application << std::endl << std::endl;
                options.output->flush();
                break;
            }
            case xlsx: {
                *options.output << L"Document data:" << std::endl;
                *options.output << L"Created in " << data.created << L" by " << data.creator << std::endl;
                *options.output << L"Last modified in " << data.modified << L" by " << data.lastModifiedBy << std::endl;
                *options.output << L"Application: " << data.application << std::endl << std::endl;
                options.output->flush();
                break;
            }
        }
    }


}
