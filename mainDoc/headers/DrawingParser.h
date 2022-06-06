//
// Created by boris on 2/22/22.
//

#ifndef DOCXTOTXT_DRAWINGPARSER_H
#define DOCXTOTXT_DRAWINGPARSER_H

#include <iostream>
#include "tinyxml2.h"
#include "../ParserCommons/FileSpecificCommons/DocCommons.h"

namespace Drawing {
    using namespace std;
    using namespace tinyxml2;

    class DrawingParser {
    private:
        size_t height = 0;
        size_t width = 0;
        string lastImageId;
        options_t &options;
        docInfo_t &docInfo;
        wstringConvert convertor;
    public:
        DrawingParser(docInfo_t &docInfo, options_t &options);

        void parseDrawing(XMLElement *element);

        void insertImage();

        void flush();
    };
}

#endif //DOCXTOTXT_DRAWINGPARSER_H
