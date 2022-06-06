//
// Created by borzor on 12/14/22.
//

#ifndef DOCXTOTXT_XLSCOMMONS_H
#define DOCXTOTXT_XLSCOMMONS_H

#include "../DocumentCommons.h"

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

struct sheet {
    std::vector<std::vector<sheetCell>> sheetArray;
    std::vector<columnSettings> col;
    std::wstring sheetName;
    std::wstring state;
    std::string docName;
    relations_t relations;

//    bool operator==(const std::string &str) const {
//        return (docName == str);
//    }
};

typedef struct {
    std::vector<std::wstring> sharedStrings;
    std::vector<draw> draws;
    std::vector<sheet> worksheets;
    documentData_t documentData;
//    relations_t relations;
} xlsInfo_t;
#endif //DOCXTOTXT_XLSCOMMONS_H
