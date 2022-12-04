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
#include "DocumentMetaDataCommons.h"

#ifndef DOCXTOTXT_DOCUMENTCOMMONS_H
#define DOCXTOTXT_DOCUMENTCOMMONS_H
#define CONTENT_TYPE_NAME "[Content_Types].xml"


#define DOC_MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define DOC_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"

#define XLS_STYLES_FILE "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
#define XLS_SHARED_STRINGS "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
#define XLS_WORKSHEET "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
#define XLS_WORKBOOK "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

#define PPT_MAIN_FILE "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
#define PPT_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
#define PPT_SLIDE_CONTENT_TYPE "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

#define DOC_IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#define TWIP_TO_CHARACTER 120

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
    std::wostream *output;
    uint32_t flags;
    char *filePath;
    zip_t *input;
    std::string pathToDraws;
    docTypes docType;
    bool printDocProps;
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
    std::vector<std::wstring> buffer;
    uint pointer;
} buffer;

typedef struct {
    size_t docWidth;
    size_t docHeight;
    std::map<std::string, std::string> imageRelationship;
    std::map<std::string, std::string> hyperlinkRelationship;
    buffer resultBuffer;
    stylesRelationship styles;
    defaultDocSettings defaultSettings;
    fileMetaData_t fileMetaData;
} docInfo_t;

static void addLine(buffer & resultBuffer) {
    resultBuffer.buffer.emplace_back();
    resultBuffer.pointer++;
}

typedef struct {
    size_t slideWidth;
    size_t slideHeight;
    std::map<std::string, tinyxml2::XMLDocument> slides;
    fileMetaData_t fileMetaData;
} pptInfo_t;

typedef struct {
    size_t startInd;
    size_t endIndInd;
    size_t width;
} columnSettings;

typedef struct {
    std::size_t cellNumber;
    std::string s; // ?
    std::string type;
    std::wstring text;
} sheetCell;

typedef struct {
    std::vector<std::vector<sheetCell>> sheetArray;
    std::vector<columnSettings> col;
    std::wstring sheetName;
    std::wstring state;
} sheet;

typedef struct {
    std::vector<std::wstring> sharedStrings;
    std::map<std::wstring, sheet> worksheets;
    buffer resultBuffer;
    fileMetaData_t fileMetaData;
} xlsInfo_t;


#endif //DOCXTOTXT_DOCUMENTCOMMONS_H
