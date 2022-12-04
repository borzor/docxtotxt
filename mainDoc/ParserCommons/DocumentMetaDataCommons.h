//
// Created by borzor on 10/18/22.
//

#ifndef DOCXTOTXT_DOCUMENTMETADATACOMMONS_H
#define DOCXTOTXT_DOCUMENTMETADATACOMMONS_H


#include <cstdio>
#include <string>

typedef struct {
    int slides;
    int pages;
    int words;
    int revision;
    std::wstring creator;
    std::wstring lastModifiedBy;
    std::wstring created;
    std::wstring modified;
    std::wstring application;
} fileMetaData_t;


#endif //DOCXTOTXT_DOCUMENTMETADATACOMMONS_H
