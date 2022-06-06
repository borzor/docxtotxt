//
// Created by borzor on 12/12/22.
//

#include "../headers/DocParser.h"
#include "../headers/TableParser.h"

namespace docxtotxt {
    DocParser::DocParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer) : docInfo(docInfo),
                                                                                         Parser(options, writer) {

    }

    void DocParser::parseFile(XMLDocument *mainFile) {
        auto mainElement = mainFile->RootElement()->FirstChildElement()->FirstChildElement();
        XMLElement *section = mainFile->RootElement()->FirstChildElement()->FirstChildElement("w:sectPr");
        parseSection(section);
        ParagraphParser paragraphParser(docInfo, options, writer);
        auto tableParser = TableParser(docInfo, options);
        writer.newLine();
        while (mainElement != nullptr) {
            try {
                if (!strcmp(mainElement->Value(), "w:p")) {
                    paragraphParser.parseParagraph(mainElement);
                    paragraphParser.writeResult();
                    paragraphParser.flush();
                } else if (!strcmp(mainElement->Value(), "w:tbl")) {
//                    tableParser.parseTable(mainElement);
//                    tableParser.insertTable();
//                    tableParser.flush();
                } else if (!strcmp(mainElement->Value(), "w:sdt")) {
                    auto sdtContent = mainElement->FirstChildElement("w:sdtContent");
                    if (sdtContent != nullptr) {
                        auto paragraphElement = sdtContent->FirstChildElement("w:p");
                        while (paragraphElement != nullptr) {
                            paragraphParser.parseParagraph(paragraphElement);
                            paragraphParser.writeResult();
                            paragraphParser.flush();
                            paragraphElement = paragraphElement->NextSiblingElement();
                        }
                    }
                }
                mainElement = mainElement->NextSiblingElement();
            } catch (exception &e) {
                std::cout << "Allocation failed: " << e.what() << '\n';
            }
        }
    }

    void DocParser::parseSection(XMLElement *section) {
        XMLElement *sectionProperty = section->FirstChildElement("w:pgSz");
        if (sectionProperty != nullptr) {
            if (sectionProperty->Attribute("w:w") != nullptr)
                this->docInfo.docWidth = atoi(sectionProperty->Attribute("w:w")) / TWIP_TO_CHARACTER;
            if (sectionProperty->Attribute("w:h") != nullptr)
                this->docInfo.docHeight = atoi(sectionProperty->Attribute("w:h")) / TWIP_TO_CHARACTER;
        }
    }


}
