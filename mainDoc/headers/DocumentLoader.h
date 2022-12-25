//
// Created by borzor on 11/30/22.
//
#include "tinyxml2.h"

#include "../ParserCommons/FileSpecificCommons/PptCommons.h"
#include "../ParserCommons/FileSpecificCommons/XlsCommons.h"
#include "../ParserCommons/FileSpecificCommons/DocCommons.h"
#include "../ParserCommons/CommonFunctions.h"
#include <zip.h>

#ifndef DOCXTOTXT_DOCUMENTLOADER_H
#define DOCXTOTXT_DOCUMENTLOADER_H


namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;

    /*!
	\brief Класс для загрузки и извлечения ключевой информации из файлов
    */
    class DocumentLoader {
    private:
        options_t &options;
        docInfo_t docInfo;
        pptInfo_t pptInfo;
        xlsInfo_t xlsInfo;
        map<string, string> content_types;
        wstringConvert convertor;
        tinyxml2::XMLDocument mainDoc;
        BufferWriter &writer;

        /*!
         * Загружает данные специфичные для WordprocessingML формата
         */
        void loadDocxData();

        /*!
         * Загружает данные специфичные для SpreadsheetML формата
         */
        void loadXlsxData();
        /*!
         * Загружает данные специфичные для PresentationML формата
         */
        void loadPptxData();
        /*!
         * Открывает файл и передает в указанную функцию указатель на корневой элемент
         * @param fileName Имя файла
         * @param f Функция, вызванная после открытия файла
         */
        void openFileAndParse(const string &fileName, void (DocumentLoader::*f)(XMLDocument *));
        /*!
         * Открывает файл и передает в указанную функцию указатель на корневой элемент и структуру взаимотношений
         * @param fileName Имя файла
         * @param relations Структура взаимотношений
         * @param f Функция, вызванная после открытия файла
         */
        void openFileAndParse(const string &fileName, relations_t &relations,
                              void (DocumentLoader::*f)(XMLDocument *, relations_t &));
        /*!
         * Обрабатывает элемент настроек презентации
         * @param doc Корневой элемент
         */
        void parsePresentationSettings(XMLDocument *doc);
        /*!
         * Обрабатывает элемент содержащий информацию о слайде
         * @param doc Корневой элемент
         */
        void parsePresentationSlide(XMLDocument *doc);
        /*!
         * Обрабатывает элемент содержащий информацию о заметке
         * @param doc Корневой элемент
         */
        void parseSlideNote(XMLDocument *doc);
        /*!
         * Извлекает информацию из элемента объекта в структуру текста
         * @param element Элемент объекта
         * @param object Cтруктура объекта текста
         */
        void parseSlideText(XMLElement *element, presentationText &object);
        /*!
         * Извлекает информацию из элемента объекта в структуру таблицы
         * @param element Элемент объекта
         * @param object Cтруктура объекта таблицы
         */
        void parseSlideTable(XMLElement *element, presentationTable &object);

        /*!
         * Извлекает структуру текста из элемента
         * @param element Элемент объекта
         * @return Массив структур текста
         */
        //TODO parseParagraph()
        std::vector<textBody> extractTextBody(XMLElement *element);
        /*!
         * Обрабатывает элемент Shared Strings
         * @param doc Корневой элемент
         */
        void parseSharedStrings(XMLDocument *doc);
        /*!
         * Обрабатывает элемент содержимого таблицы
         * @param doc Корневой элемент
         */
        void parseWorksheet(XMLDocument *doc);
        /*!
         * Обрабатывает элемент настроек таблиц
         * @param doc Корневой элемент
         */
        void parseWorkbook(XMLDocument *doc);
        /*! TODO
         * Обрабатывает элемент DrawingML
         * @param doc Корневой элемент
         */
        void parseDraw(XMLDocument *doc);
        /*!
         * Обрабатывает элемент [Content_Types]
         * @param doc Корневой элемент
         */
        void parseContentTypes(XMLDocument *doc);
        /*!
         * Обрабатывает элемент взаимотношений
         * @param doc Корневой элемент
         */
        void parseRelationShip(XMLDocument *doc, relations_t &relations);
        /*!
         * Обрабатывает элемент стилей
         * @param doc Корневой элемент
         */
        void parseStyles(XMLDocument *doc);
        /*!
         * Добавляет элемент содержащий описание стиля в списое стилей
         * @param doc Корневой элемент
         */
        void addStyle(XMLElement *element);
        /*!
         * Устанавливает дефолтные стили
         * @param doc Корневой элемент
         */
        void setDefaultSettings(XMLElement *element);
        /*!
         * Обрабатывает App файл, содержащий мета информацию документа
         * @param doc Корневой элемент
         */
        void parseAppFile(XMLDocument *doc);
        /*!
         * Обрабатывает Core файл, содержащий мета информацию документа
         * @param doc Корневой элемент
         */
        void parseCoreFile(XMLDocument *doc);
        /*!
         * Обрабатывает основной файл WordprocessingML
         * @param doc Корневой элемент
         */
        void parseDocFile(XMLDocument *doc);
        /*!
         * Обрабатывает параграф и заполняет переданную структуру
         * @param paragraph Элемент параграфа
         * @param par Структура параграфа
         */
        void parseParagraph(XMLElement *paragraph, ::docxtotxt::paragraph &par);
        /*!
         * Обабрабатывает текстовое содержание параграфа и заполняет переданную структуру
         * @param properties Элемент текстового содержания параграфа
         * @param paragraph Структура параграфа
         */
        void parseTextProperties(XMLElement *properties, paragraph &paragraph);
        /*!
         * Обрабатывает таблицу и заполняет переданную структуру
         * @param paragraph Элемент параграфа
         * @param par Структура параграфа
         */
        void parseTable(XMLElement *paragraph, ::docxtotxt::paragraph &par);
        /*!
         * Обрабатывает раздел секции и устанавливает соответствующие настройки
         * @param section Элемент секции
         */
        void parseSection(XMLElement *section);


    public:
        explicit DocumentLoader(options_t &options, BufferWriter &writer);

        /*!
         * Разархивирует входные данные и извлекает ключевую информацию на основе предоставленных параметров
         */
        void loadData();

        /*!
         * Возвращает структуру WordprocessingML документа
         * @return Структура WordprocessingML документа
         */
        docInfo_t getDocxData() const;

        /*!
         * Возвращает структуру SpreadsheetML документа
         * @return Структура SpreadsheetML документа
         */
        xlsInfo_t getXlsxData() const;

        /*!
         * Возвращает структуру PresentationML документа
         * @return Структура PresentationML документа
         */
        pptInfo_t getPptxData() const;
    };

}
#endif //DOCXTOTXT_DOCUMENTLOADER_H
