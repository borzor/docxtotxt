//
// Created by borzor on 10/18/22.
//

#ifndef DOCXTOTXT_DOCUMENTMETADATACOMMONS_H
#define DOCXTOTXT_DOCUMENTMETADATACOMMONS_H


#include <cstdio>
#include <string>

namespace docxtotxt {
    typedef struct {
        size_t slides;
        size_t pages;
        size_t words;
        size_t revision;
        std::wstring creator;
        std::wstring lastModifiedBy;
        std::wstring created;
        std::wstring modified;
        std::wstring application;
    } fileMetadata_t;
}

#endif //DOCXTOTXT_DOCUMENTMETADATACOMMONS_H
