
//
// Created by borzor on 12/14/22.
//

#ifndef DOCXTOTXT_DOCCOMMONS_H
#define DOCXTOTXT_DOCCOMMONS_H

#include "../DocumentCommons.h"
#include "../BufferWriter.h"

namespace docxtotxt {

    typedef struct {//TODO add rest doc styles
        paragraphSettings paragraph;
        tableProperties table;
    } defaultDocStyles;

    typedef struct {
        std::map<std::string, paragraphSettings> paragraphStyles;
        std::map<std::string, tableProperties> tableStyles;
        defaultDocStyles defaultStyles;
    } stylesRelationship;

    typedef struct {
        paragraphSettings paragraph;
    } defaultDocSettings;

    typedef struct {
        size_t docWidth;
        size_t docHeight;
        relations_t relations;
        stylesRelationship styles;
        defaultDocSettings defaultSettings;
    } docInfo_t;
}
#endif //DOCXTOTXT_DOCCOMMONS_H
