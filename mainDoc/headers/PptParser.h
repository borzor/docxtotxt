//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_PPTPARSER_H
#define DOCXTOTXT_PPTPARSER_H

#include "../ParserCommons/FileSpecificCommons/PptCommons.h"

#include "Parser.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;
    /*!
	\brief Реализация парсера, обрабатывающая PresentationML формат
    */
    class PptParser : public Parser{

    private:
        ///Настройки специфичные для файла PresentationML формата
        pptInfo_t &pptInfo;
        std::vector<slideInsertInfo> slideInsertData;
        /*!
         * Преобразовывает входные структуры в структуру для вставки
         */
        void prepareSlides();
        /*!
         * Записывает объекты слайда
         * @param insertObjects Массив вставляемых объектов
         */
        void insertSlide(std::vector<insertObject> insertObjects);
        /*!
         * Записывает часть обрабатываемого объекта в результируюий буфер
         * @param obj Структура записываемого объекта
         * @param currentIndex Текущий индекс, необходимый для установки размера вставки
         */
        void insertLineBuffer(insertObject &obj, size_t *currentIndex);
        /*!
         * Зписывает мета информацию о слайде в результирующий буфер
         * @param relations Структура метаданных
         */
        void insertSlideMetadata(slideInsertInfo relations);

    public:
        PptParser(pptInfo_t &pptInfo, options_t &options, BufferWriter &writer);

        void parseFile() override;
    };
}

#endif //DOCXTOTXT_PPTPARSER_H
