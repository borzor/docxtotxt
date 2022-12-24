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
         * test
         */
        void loadDocxData();

        /*!
         * test
         */
        void loadXlsxData();

        void loadPptxData();

        void openFileAndParse(const string &fileName, void (DocumentLoader::*f)(XMLDocument *));

        void openFileAndParse(const string &fileName, relations_t &relations,
                              void (DocumentLoader::*f)(XMLDocument *, relations_t &));

        void parsePresentationSettings(XMLDocument *doc);

        void parsePresentationSlide(XMLDocument *doc);

        void parseSlideNote(XMLDocument *element);

        void parseSlideText(XMLElement *element, presentationText &object);

        void parseSlideTable(XMLElement *element, presentationTable &object);

        std::vector<textBody> extractTextBody(XMLElement *element);

        void parseSharedStrings(XMLDocument *doc);

        void parseWorksheet(XMLDocument *doc);

        void parseWorkbook(XMLDocument *doc);

        void parseDraw(XMLDocument *doc);

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseRelationShip(XMLDocument *doc, relations_t &relations);

        void parseStyles(XMLDocument *doc);

        void addStyle(XMLElement *element);

        void setDefaultSettings(XMLElement *element);

        void parseAppFile(XMLDocument *doc);

        void parseCoreFile(XMLDocument *doc);

        void parseDocFile(XMLDocument *doc);

        void parseParagraph(XMLElement *paragraph, ::docxtotxt::paragraph &par);

        void parseTextProperties(XMLElement *properties, paragraph &paragraph);

        void parseTable(XMLElement *paragraph, ::docxtotxt::paragraph &par);

        void parseSection(XMLElement *section);


    public:
        explicit DocumentLoader(options_t &options, BufferWriter &writer);

        /*!
         * Разархивирует входные данные и извлекает ключевую информацию на основе предоставленных параметров
         */
        void loadData();

        /*!
         * Возвращает структуру .docx документа
         * @return Структура .docx документа
         */
        docInfo_t getDocxData() const;

        /*!
         * Возвращает структуру .xlsx документа
         * @return Структура .xlsx документа
         */
        xlsInfo_t getXlsxData() const;

        /*!
         * Возвращает структуру .pptx документа
         * @return Структура .pptx документа
         */
        pptInfo_t getPptxData() const;
    };

}
#endif //DOCXTOTXT_DOCUMENTLOADER_H
