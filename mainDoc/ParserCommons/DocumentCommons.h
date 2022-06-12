//
// Created by boris on 6/9/22.
//
#include "ParagraphCommons.h"

#ifndef DOCXTOTXT_DOCUMENTCOMMONS_H
#define DOCXTOTXT_DOCUMENTCOMMONS_H

#define MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
#define IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#define TWIP_TO_CHARACTER 120

using wstringConvert = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>;

inline bool ends_with(std::string const &value, std::string const &ending) {
    if (ending.size() > value.size()) return false;
    return std::equal(ending.rbegin(), ending.rend(), value.rbegin());
}

typedef struct {
    uint32_t flags;
    zip_t *input;
    std::wostream *output;
    std::string pathToDraws;
} options_t;

typedef struct {//TODO add rest doc styles
    paragraphSettings paragraph;
} defaultDocStyles;

typedef struct {
    std::map<std::string, paragraphSettings> paragraphStyles;
//    std::map<std::string, std::string> tableStyles;
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
    uint pointer;
    std::map<std::string, std::string> imageRelationship;
    std::map<std::string, std::string> hyperlinkRelationship;
    std::vector<std::pair<std::wstring, size_t>> docBuffer;
    stylesRelationship styles;
    defaultDocSettings defaultSettings;
} docInfo_t;


#endif //DOCXTOTXT_DOCUMENTCOMMONS_H
