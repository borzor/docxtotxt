//
// Created by boris on 2/21/22.
//

#ifndef DOCXTOTXT_SECTIONPARSER_H
#define DOCXTOTXT_SECTIONPARSER_H
#include <iostream>
#include "tinyxml2.h"
#include "tinyxml.h"
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
    private://TODO default value
        size_t height = 16838; // default value, in TWIP's
        size_t width = 11906;

    public:
        void parseSection(XMLElement *section);

        size_t getDocWidth();
    };
}

#endif //DOCXTOTXT_SECTIONPARSER_H
