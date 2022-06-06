
//
// Created by borzor on 12/14/22.
//

#ifndef DOCXTOTXT_DOCCOMMONS_H
#define DOCXTOTXT_DOCCOMMONS_H

#include "../DocumentCommons.h"

typedef struct {//TODO add rest doc styles
    paragraphSettings paragraph;
    tableProperties table;
} defaultDocStyles;

typedef struct {
    std::map<std::string, paragraphSettings> paragraphStyles;
    std::map<std::string, tableProperties> tableStyles;
//    std::map<std::string, std::string> characterStyles;
//    std::map<std::string, std::string> numberingStyles;
    defaultDocStyles defaultStyles;
} stylesRelationship;

typedef struct {
    paragraphSettings paragraph;
    textSettings text;
} defaultDocSettings;

typedef struct {
    size_t docWidth;
    size_t docHeight;
    documentData_t documentData;
    relations_t relations;
    stylesRelationship styles;
    defaultDocSettings defaultSettings;
} docInfo_t;

#endif //DOCXTOTXT_DOCCOMMONS_H
