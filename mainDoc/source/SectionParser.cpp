//
// Created by boris on 2/21/22.
//

#include "../headers/SectionParser.h"

namespace section {
    void SectionParser::parseSection(XMLElement *section) {
        XMLElement *sectionProperty = section->FirstChildElement();
        while (sectionProperty != nullptr) {
            switch (sectionProperties[sectionProperty->Value()]) {
                case cols:
                    break;
                case footerReference:
                    break;
                case formProt:
                    break;
                case headerReference:
                    break;
                case lnNumType:
                    break;
                case paperSrc:
                    break;
                case pgBorders:
                    break;
                case pgMar:
                    break;
                case pgNumType:
                    break;
                case pgSz:
                    if(sectionProperty->Attribute("w:w") != nullptr)
                        this->docInfo.docWidth = atoi(sectionProperty->Attribute("w:w"))/ TWIP_TO_CHARACTER;
                    if(sectionProperty->Attribute("w:h") != nullptr)
                        this->docInfo.docHeight = atoi(sectionProperty->Attribute("w:h"))/ TWIP_TO_CHARACTER;
                    break;
                case titlePg:
                    break;
                case type:
                    break;
                case vAlign:
                    break;
            }
            sectionProperty = sectionProperty->NextSiblingElement();
        }
        free(sectionProperty);
    }

    size_t SectionParser::getDocWidth() {
        return this->docInfo.docWidth;
    }

    size_t SectionParser::getDocHeight() {
        return this->docInfo.docHeight;
    }

    SectionParser::SectionParser(::docInfo_t &docInfo) : docInfo(docInfo) {

    }

}
