//
// Created by borzor on 12/14/22.
//

#ifndef DOCXTOTXT_PPTCOMMONS_H
#define DOCXTOTXT_PPTCOMMONS_H

#include "../DocumentCommons.h"

typedef struct {
    size_t slideWidth;
    size_t slideHeight;
    size_t widthCoefficient;
    size_t heightCoefficient;
} presentationSettings;

enum algn {
    l,
    r,
    ctr
};
typedef struct {
    algn align;
    std::wstring text;
} textBody;

typedef struct {
    size_t offsetX;
    size_t offsetY;
    size_t objectSizeX;
    size_t objectSizeY;
} objectInfo_t;

typedef struct {
    objectInfo_t objectInfo;
    std::vector<size_t> gridColSize;
    std::vector<std::vector<std::vector<textBody>>> table;
} presentationTable;

typedef struct {
    objectInfo_t objectInfo;
    std::vector<textBody> paragraph;
} presentationText;

typedef struct {
    objectInfo_t objectInfo;
    std::string rId;
    std::string description;
} presentationPic;

struct insertObject {
    size_t startLine;
    size_t offset;
    size_t length;
    bool inProgress;
    std::vector<textBody> paragraph;

    bool operator<(const insertObject &str) const {
        return (offset < str.offset);
    }
};

struct insertImage {
    size_t startLine;
    size_t offset;
    size_t length;
    std::string rId;

    bool operator==(const std::string &str) const {
        return (rId == str);
    }

    std::wstring toString() const {
        return L"Line - " + std::to_wstring(startLine) + L" Offset - " + std::to_wstring(offset) + L" Length - " +
               std::to_wstring(length);
    }
};

typedef struct {
    std::vector<insertObject> insertObjects;
    std::vector<insertImage> insertImages;
    relations_t relations;
} slideInsertInfo;

struct noteInfo_t {
    std::string noteName;
    std::vector<textBody> text;

    bool operator==(const std::string &str) const {
        return (noteName == str);
    }
};

typedef struct {
    std::vector<presentationText> objects;
    std::vector<presentationTable> tables;
    std::vector<presentationPic> pictures;
    relations_t relations;
} slideInfo;

typedef struct {
    presentationSettings settings;
    std::vector<slideInfo> slides;
    documentData_t documentData;
    std::vector<noteInfo_t> notes;
} pptInfo_t;

#endif //DOCXTOTXT_PPTCOMMONS_H
