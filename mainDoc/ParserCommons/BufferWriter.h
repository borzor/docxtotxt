//
// Created by borzor on 12/17/22.
//
#include <iostream>
#include "../ParserCommons/DocumentCommons.h"

#ifndef DOCXTOTXT_BUFFERWRITER_H
#define DOCXTOTXT_BUFFERWRITER_H

namespace docxtotxt {

    class BufferWriter {
    private:
        std::vector<std::wstring> buffer;
        uint pointer = 0;
        fileMetadata_t fileMetadata;
        wstringConvert convertor;
        options_t options;

    public:
        explicit BufferWriter(options_t &options);

        std::wstring convertString(const std::string &str);

        void newLine();

        void insertData(const std::wstring &data, bool addLineBefore = false, bool addLineAfter = true);

        void insertMetadata();

        void writeResult();

        fileMetadata_t *getMetadata();

        size_t getCurrentLength() const;

    };

} // docxtotxt

#endif //DOCXTOTXT_BUFFERWRITER_H
