//
// Created by borzor on 12/12/22.
//

#ifndef DOCXTOTXT_DOCPARSER_H
#define DOCXTOTXT_DOCPARSER_H

#include "../ParserCommons/FileSpecificCommons/DocCommons.h"
#include "Parser.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;
    /*!
	\brief Реализация парсера, обрабатывающая WordprocessingML формат
    */
    class DocParser : public Parser {
    private:
        ///Настройки специфичные для файла WordprocessingML формата
        docInfo_t &docInfo;
        /*!
         * Записывает параграф в результирующий буфер с учетом форматирования
         * @param paragraph Структура параграфа
         */
        void writeParagraph(paragraph paragraph);
        /*!
         * Записывает параграф содержащий изображение в результирующий буфер с учетом форматирования
         * @param paragraph Структура параграфа
         */
        void writeImage(paragraph paragraph);
        /*!
         * Записывает параграф содержащий таблицу в результирующий буфер с учетом форматирования
         * @param paragraph Структура параграфа
         */
        void writeTable(paragraph paragraph);
    public:
        DocParser(docInfo_t &docInfo, options_t &options, BufferWriter &writer);

        void parseFile() override;

    };

}
#endif //DOCXTOTXT_DOCPARSER_H
