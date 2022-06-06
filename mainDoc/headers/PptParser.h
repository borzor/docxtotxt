//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_PPTPARSER_H
#define DOCXTOTXT_PPTPARSER_H

#include "../ParserCommons/FileSpecificCommons/PptCommons.h"

namespace ppt {
    using namespace std;
    using namespace tinyxml2;

    class PptParser {

    private:
        options_t &options;
        pptInfo_t &pptInfo;
        std::vector<slideInsertInfo> slideInsertData;
        wstringConvert convertor;

        void prepareSlides();

        void insertSlide(std::vector<insertObject> insertObjects);

        void insertLineBuffer(insertObject &obj, size_t *currentIndex);

        void insertSlideMetaData(slideInsertInfo relations);

    public:
        PptParser(pptInfo_t &pptInfo, options_t &options);

        void parseSlides();
    };
}

#endif //DOCXTOTXT_PPTPARSER_H
