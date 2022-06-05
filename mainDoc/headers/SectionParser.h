//
// Created by boris on 2/21/22.
//

#ifndef DOCXTOTXT_SECTIONPARSER_H
#define DOCXTOTXT_SECTIONPARSER_H
#include <iostream>
#include "tinyxml2.h"
#include "tinyxml.h"
#include "ParagraphParser.h"
#include <map>

#define TWIP_TO_CHARACTER 120

namespace section {
    using namespace std;
    using namespace tinyxml2;

    enum sectionProperty {
        cols,
        footerReference,
        formProt,
        headerReference,
        lnNumType,
        paperSrc,
        pgBorders,
        pgMar,
        pgNumType,
        pgSz,
        titlePg,
        type,
        vAlign
    };
    static map<string, int> sectionProperties = {{"w:cols",            cols},
                                                 {"w:footerReference", footerReference},
                                                 {"w:formProt",        formProt},
                                                 {"w:headerReference", headerReference},
                                                 {"w:lnNumType",       lnNumType},
                                                 {"w:paperSrc",        paperSrc},
                                                 {"w:pgBorders",       pgBorders},
                                                 {"w:pgMar",           pgMar},
                                                 {"w:pgNumType",       pgNumType},
                                                 {"w:pgSz",            pgSz},
                                                 {"w:titlePg",         titlePg},
                                                 {"w:type",            type},
                                                 {"w:vAlign",          vAlign}};

    class SectionParser {
    private:
        docInfo_t &docInfo;
    public:
        explicit SectionParser(::docInfo_t &docInfo);

        void parseSection(XMLElement *section);

        size_t getDocWidth();

        size_t getDocHeight();
    };
}

#endif //DOCXTOTXT_SECTIONPARSER_H
