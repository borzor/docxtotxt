//
// Created by boris on 6/9/22.
//
#include <locale>
#include <codecvt>
#include <zip.h>
#include <map>
#include <tinyxml2.h>
#include "DocumentMetaDataCommons.h"
#include <algorithm>
#include <vector>

#ifndef DOCXTOTXT_DOCUMENTCOMMONS_H
#define DOCXTOTXT_DOCUMENTCOMMONS_H
namespace docxtotxt {
#define CONTENT_TYPE_NAME "[Content_Types].xml"

#define DOC_MAIN_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
#define DOC_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
#define DOC_IMAGE_FILE_PATH "word/_rels/document.xml.rels"

#define XLS_STYLES_FILE "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
#define XLS_SHARED_STRINGS "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
#define XLS_WORKSHEET "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
#define XLS_WORKBOOK "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
#define XLS_SLIDE_NOTE "application/vnd.openxmlformats-officedocument.drawing+xml"

#define PPT_STYLES_FILE "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
#define PPT_MAIN_FILE "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
#define PPT_SLIDE_CONTENT_TYPE "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
#define PPT_SLIDE_NOTE "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
#define PPT_TABLE_URI "http://schemas.openxmlformats.org/drawingml/2006/table"

#define TWIP_TO_CHARACTER 120

#define PRESENTATION_WIDTH 150
#define PRESENTATION_HEIGHT 40

    using wstringConvert = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>;

    inline bool ends_with(std::string const &value, std::string const &ending) {
        if (ending.size() > value.size()) return false;
        return std::equal(ending.rbegin(), ending.rend(), value.rbegin());
    }

    inline bool starts_with(std::string const &value, std::string const &starting) {
        return value.rfind(starting, 0) == 0;
    }

    enum justify_t {
        left,
        right,
        center,
        both,
        distribute
    };

    enum tabCharacter {
        dot,
        heavy,
        hyphen,
        middleDot,
        none,
        underscore
    };

    typedef struct {
        int left;
        int right;
        int hanging;
        int firstLine;
    } indentation;

    typedef struct {
        size_t before;
        size_t after;
    } lineSpacing;

    typedef struct { //val needed?
        size_t pos;
        tabCharacter character;
    } tabulation;

    typedef struct {
        std::map<std::string, std::string> imageRelationship;
        std::map<std::string, std::string> hyperlinkRelationship;
        std::map<std::string, std::string> notes;
        std::map<std::string, std::string> drawing;
    } relations_t;
    typedef struct {
        size_t offsetX;
        size_t offsetY;
        size_t objectSizeX;
        size_t objectSizeY;
    } objectInfo_t;

    typedef struct {
        objectInfo_t objectInfo;
        std::string rId;
        std::string description;
    } picture;

    struct draw {
        std::string name;
        picture pic;
        relations_t relations;

        bool operator==(const std::string &str) const {
            return (name == str);
        }
    };

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
    } options_t;

}
#endif //DOCXTOTXT_DOCUMENTCOMMONS_H
