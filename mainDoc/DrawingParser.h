//
// Created by boris on 2/22/22.
//

#ifndef DOCXTOTXT_DRAWINGPARSER_H
#define DOCXTOTXT_DRAWINGPARSER_H

#include <iostream>
#include "tinyxml2.h"
#include "tinyxml.h"

namespace Drawing {
    using namespace std;
    using namespace tinyxml2;

    class DrawingParser {
    private:
        const XMLElement *drawing;
        size_t height = 0;
        size_t width = 0;
    public:
        explicit DrawingParser(XMLElement *drawing);

        void parseDrawing();

        void getDrawingSize(size_t &height, size_t &width);
    };
}

#endif //DOCXTOTXT_DRAWINGPARSER_H
