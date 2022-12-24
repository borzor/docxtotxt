
//
// Created by borzor on 12/14/22.
//

#ifndef DOCXTOTXT_DOCCOMMONS_H
#define DOCXTOTXT_DOCCOMMONS_H

#include "../DocumentCommons.h"
#include "../BufferWriter.h"

namespace docxtotxt {

    typedef struct {
        bool outline;
        justify_t justify;
        indentation ind;
        lineSpacing spacing;
        std::vector<tabulation> tab;
    } paragraphSettings;
    enum class tableJustify {
        start,
        end,
        center
    };
    typedef struct {
        tableJustify justify;
        size_t ind;
    } tableProperties;

    typedef struct {
        size_t width;
        size_t gridSpan;
        std::wstring text;
    } cell;

    typedef struct {
        size_t columnAmount;
        tableProperties settings;
        std::vector<std::size_t> tblGrids;
        std::vector<std::vector<cell>> grid;
    } tableSetting;

    enum paragraphType {
        par,
        table,
        image
    };
    typedef struct {
        paragraphSettings settings;
        std::vector<std::wstring> body;
        paragraphType type;
    } paragraph;

    typedef struct {
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
        std::vector<paragraph> body;
    } docInfo_t;


}
#endif //DOCXTOTXT_DOCCOMMONS_H
