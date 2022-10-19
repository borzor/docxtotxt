//
// Created by boris on 6/9/22.
//
#include <locale>
#include <codecvt>
#include <zip.h>
#include <map>
#include <tinyxml2.h>
#include "ParagraphCommons.h"
#include "TableCommons.h"
#ifndef DOCXTOTXT_DOCUMENTCOMMONS_H
#define DOCXTOTXT_DOCUMENTCOMMONS_H

#define DOC_MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define DOC_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"

#define PPT_MAIN_FILE "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
#define PPT_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
#define PPT_SLIDE_CONTENT_TYPE "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

#define DOC_IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#define TWIP_TO_CHARACTER 120
#define TWIP_TO_CHARACTER_HEIGHT 240
using wstringConvert = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>;

inline bool ends_with(std::string const &value, std::string const &ending) {
    if (ending.size() > value.size()) return false;
    return std::equal(ending.rbegin(), ending.rend(), value.rbegin());
}

inline bool starts_with(std::string const &value, std::string const &starting) {
    return value.rfind(starting, 0) == 0;
}

enum docTypes {
    pptx,
    docx,
    xlsx
};
typedef struct {
    uint32_t flags;
    zip_t *input;
    std::wostream *output;
    std::string pathToDraws;
    docTypes docType;
} options_t;

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
    uint pointer;
    std::map<std::string, std::string> imageRelationship;
    std::map<std::string, std::string> hyperlinkRelationship;
    std::vector<std::pair<std::wstring, size_t>> docBuffer;
    stylesRelationship styles;
    defaultDocSettings defaultSettings;
} docInfo_t;

typedef struct {
    size_t slideWidth;
    size_t slideHeight;
    std::map<std::string, tinyxml2::XMLDocument> slides;
} pptInfo_t;


#endif //DOCXTOTXT_DOCUMENTCOMMONS_H
