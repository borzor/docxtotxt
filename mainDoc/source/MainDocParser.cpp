//
// Created by boris on 2/9/22.
//

#include <string>
#include <locale>
#include <fstream>
#include "../headers/MainDocParser.h"
#include "../headers/XlsParser.h"
#include "../headers/DocumentLoader.h"
#include "../headers/PptParser.h"
#include "../headers/DocParser.h"

namespace docxtotxt {
    void MainDocParser::parseFile() {
        docxtotxt::BufferWriter writer(options);
        DocumentLoader documentLoader(options, writer);
        documentLoader.loadData();
        if ((options.flags >> 4) & 1) {
            writer.insertMetadata();
        }
        switch (options.docType) {
            case pptx: {
                auto pptInfo = documentLoader.getPptxData();
                PptParser pptParser(pptInfo, options,writer);
                pptParser.parseFile();
                if ((options.flags >> 1) & 1) {
                    for (auto &slide: pptInfo.slides)
                        saveImages(slide.relations.imageRelationship);
                }
                writer.writeResult();
                break;
            }
            case docx: {
                auto docInfo = documentLoader.getDocxData();
                DocParser docParser(docInfo, options, writer);
                docParser.parseFile();
                if ((options.flags >> 1) & 1) {
                    saveImages(docInfo.relations.imageRelationship);
                }
                writer.writeResult();
                if ((options.flags >> 2) & 1) {
                    insertHyperlinks(docInfo.relations.hyperlinkRelationship);
                }
                break;
            }
            case xlsx: {
                auto xlsInfo = documentLoader.getXlsxData();
                options.output->flush();
                XlsParser xlsParser(xlsInfo, options, writer);
                xlsParser.parseFile();
                if ((options.flags >> 1) & 1) {
                    for (auto &draw: xlsInfo.draws)
                        saveImages(draw.relations.imageRelationship);
                }
                writer.writeResult();
                break;
            }
        }
        zip_close(options.input);
        delete (options.output);
    }

    MainDocParser::MainDocParser(options_t &options) :
            options(options) {

    }

    void MainDocParser::insertHyperlinks(std::map<std::string, std::string> hyperlinkRelationship) {
        for (const auto &kv: hyperlinkRelationship) {
            auto number = distance(hyperlinkRelationship.begin(),
                                   hyperlinkRelationship.find(kv.first));
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
            if (zip_stat(options.input, kv.second.c_str(), ZIP_FL_NODIR, &file_info) == -1)throw runtime_error("Error: Cannot get info about " + kv.second + " file");
            char tmpImageBuffer[file_info.size];
            auto currentImage = zip_fopen(options.input, kv.second.c_str(), ZIP_FL_NODIR);
            if (zip_fread(currentImage, &tmpImageBuffer, file_info.size) == -1)throw runtime_error("Error: Cannot read " + kv.second + " file");
            zip_fclose(currentImage);
            if (!std::ofstream(string(options.pathToDraws) + "/" + kv.second).write(tmpImageBuffer, file_info.size)) {throw runtime_error("Error writing file" + kv.second);
            }
        }
    }







}
