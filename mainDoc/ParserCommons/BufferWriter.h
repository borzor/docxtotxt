//
// Created by borzor on 12/17/22.
//
#include <iostream>
#include "../ParserCommons/DocumentCommons.h"

#ifndef DOCXTOTXT_BUFFERWRITER_H
#define DOCXTOTXT_BUFFERWRITER_H

namespace docxtotxt {
    /*!
	\brief Класс для изменения внутреннего буфера и записи результата в поток вывода
    */
    class BufferWriter {
    private:
        std::vector<std::wstring> buffer;
        uint pointer = 0;
        fileMetadata_t fileMetadata;
        ///Конвертор string в wstring
        wstringConvert convertor;
        options_t options;

    public:
        explicit BufferWriter(options_t options);

        /*!
         * Преобразовывает входную строку string и возвращает wstring
         * @param str Входная строка
         * @return Результат
         */
        std::wstring convertString(const std::string &str);

        /*!
         * Добавляет новую строку в результирующий буфер
         */
        void newLine();

        /*!
         * Добавляет текст к результирующему буферу
         * @param data Текст которую необходимо добавить
         * @param addLineBefore Флаг для обозначения необходимости вставить новую строку перед вставкой текста
         * @param addLineAfter Флаг для обозначения необходимости вставить новую строку после вставки текста
         */
        void insertData(const std::wstring &data, bool addLineBefore = false, bool addLineAfter = true);

        /*!
         * Вставляет метаданные документа
         */
        void insertMetadata();

        /*!
         * Записывает результат в поток вывода
         */
        void writeResult();

        /*!
         * Возвращает указатель на структуру метаданных документа
         * @return Указатель на структуру метаданных документа
         */
        fileMetadata_t *getMetadata();

        /*!
         * Возвращает длину текущей строки буфера
         * @return Длина текущей строки буфера
         */
        size_t getCurrentLength() const;

    };

} // docxtotxt

#endif //DOCXTOTXT_BUFFERWRITER_H
