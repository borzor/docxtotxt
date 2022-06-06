//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_PPTPARSER_H
#define DOCXTOTXT_PPTPARSER_H

#include "../ParserCommons/FileSpecificCommons/PptCommons.h"
#include "../ParserCommons/BufferWriter.h"
#include "Parser.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;
    /*!
	\brief Реализация парсера, обрабатывающая .pptx формат
    */
    class PptParser : public Parser{

    private:
        ///Настройки специфичные для файла .pptx формата
        pptInfo_t &pptInfo;
        std::vector<slideInsertInfo> slideInsertData;

        void prepareSlides();

        void insertSlide(std::vector<insertObject> insertObjects);

        void insertLineBuffer(insertObject &obj, size_t *currentIndex);

        void insertSlideMetadata(slideInsertInfo relations);

    public:
        PptParser(pptInfo_t &pptInfo, options_t &options, BufferWriter &writer);

        void parseFile() override;
    };
}

#endif //DOCXTOTXT_PPTPARSER_H
