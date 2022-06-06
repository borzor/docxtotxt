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

typedef struct {
    std::vector<std::vector<sheetCell>> sheetArray;
    std::vector<columnSettings> col;
    std::wstring sheetName;
    std::wstring state;
} sheet;

typedef struct {
    std::vector<std::wstring> sharedStrings;
    std::map<std::wstring, sheet> worksheets;
    documentData_t documentData;
    relations_t relations; //todo
} xlsInfo_t;
#endif //DOCXTOTXT_XLSCOMMONS_H
