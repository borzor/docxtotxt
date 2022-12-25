//
// Created by borzor on 12/17/22.
//

#include <string>
#include "../ParserCommons/BufferWriter.h"

namespace docxtotxt {
    BufferWriter::BufferWriter(options_t &options) : options(options) {
        buffer.emplace_back();
        fileMetadata.slides = 0;
    }

    void BufferWriter::newLine() {
        buffer.emplace_back();
        pointer++;
    }

    void BufferWriter::insertData(const std::wstring &data, bool addLineBefore, bool addLineAfter) {
        if (addLineBefore)
            newLine();
        buffer.back().append(data);
        if (addLineAfter)
            newLine();
    }

    void BufferWriter::writeResult() {
        for (const auto &kv: buffer) {
            *options.output << kv << std::endl;
        }
    }

    std::wstring BufferWriter::convertString(const std::string &str) {
        return convertor.from_bytes(str);
    }

    void BufferWriter::insertMetadata() {
        insertData(L"Document data:");
        insertData(L"Application: " + fileMetadata.application);
        insertData(L"Created in " + fileMetadata.created + L" by " + fileMetadata.creator);
        insertData(L"Last modified in " + fileMetadata.modified + L" by " + fileMetadata.lastModifiedBy);
        switch (options.docType) {
            case pptx: {
                insertData(L"Revision №" + std::to_wstring(fileMetadata.revision) + L", contains " +
                           std::to_wstring(fileMetadata.slides) + L" slides and " +
                           std::to_wstring(fileMetadata.words) + L" words");
                break;
            }
            case docx: {
                insertData(L"Revision №" + std::to_wstring(fileMetadata.revision) + L", contains " +
                           std::to_wstring(fileMetadata.pages) + L" pages and " +
                           std::to_wstring(fileMetadata.words) + L" words");
                break;
            }
            case xlsx: {
                break;
            }
        }

    }

    fileMetadata_t *BufferWriter::getMetadata() {
        return &fileMetadata;
    }

    size_t BufferWriter::getCurrentLength() const {
        return buffer.back().length();
    }


} // docxtotxt