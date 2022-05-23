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
        size_t height = 0;
        size_t width = 0;
        string lastImageId;
    public:
        explicit DrawingParser(bool saveDraws);

        void parseDrawing(XMLElement *element);

        void getDrawingSize(size_t &height, size_t &width) const;

        string getDrawingId() const;
    };
}

#endif //DOCXTOTXT_DRAWINGPARSER_H
