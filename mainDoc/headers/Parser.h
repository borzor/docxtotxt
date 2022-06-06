//
// Created by borzor on 12/17/22.
//

#ifndef DOCXTOTXT_PARSER_H
#define DOCXTOTXT_PARSER_H

#include "../ParserCommons/BufferWriter.h"

namespace docxtotxt {
    /*!
	\brief Интерфейс парсера
    */
    class Parser {
    public:
        explicit Parser(options_t &options, BufferWriter &writer) : options(options), writer(writer) {}

        /*!
         * Функия для обработки входного файла на основе предоставленных опций
         */
        virtual void parseFile() = 0;

    protected:
        ///Структура настроек, уставляемых во время инициализации программы
        options_t &options;
        ///Класс реализующий запись в результирующий буфер
        BufferWriter &writer;

    };

}


#endif //DOCXTOTXT_PARSER_H
