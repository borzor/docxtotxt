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
                        this->width = atoi(sectionProperty->Attribute("w:w"));
                    if(sectionProperty->Attribute("w:h") != nullptr)
                        this->height = atoi(sectionProperty->Attribute("w:h"));
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
        return this->width / TWIP_TO_CHARACTER;
    }

    size_t SectionParser::getDocHeight() {
        return this->height / TWIP_TO_CHARACTER;
    }

}
