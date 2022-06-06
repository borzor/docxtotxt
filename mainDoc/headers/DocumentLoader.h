//
// Created by borzor on 11/30/22.
//
#include "tinyxml2.h"

#include "../ParserCommons/FileSpecificCommons/PptCommons.h"
#include "../ParserCommons/FileSpecificCommons/XlsCommons.h"
#include "../ParserCommons/FileSpecificCommons/DocCommons.h"
#include <zip.h>

#ifndef DOCXTOTXT_DOCUMENTLOADER_H
#define DOCXTOTXT_DOCUMENTLOADER_H

namespace Loader {
    using namespace std;
    using namespace tinyxml2;

    class DocumentLoader {
    private:
        options_t &options;
        docInfo_t docInfo;
        pptInfo_t pptInfo;
        xlsInfo_t xlsInfo;
        map<string, string> content_types;
        wstringConvert convertor;
        tinyxml2::XMLDocument mainDoc;

        void loadDocxData();

        void loadXlsxData();

        void loadPptxData();

        void openFileAndParse(const string &fileName, void (DocumentLoader::*f)(XMLDocument *));

        void parsePresentationSettings(XMLDocument *doc);

        void parsePresentationSlide(XMLDocument *doc);

        void parseSlideNote(XMLDocument *element);

        void parseSlideText(XMLElement *element, presentationText &object);

        void parseSlideTable(XMLElement *element, presentationTable &object);

        void parseSlidePic(XMLElement *element, presentationPic &object);

        std::vector<textBody> extractTextBody(XMLElement *element);

        objectInfo_t extractObjectInfo(XMLElement *element);

        void parseSharedStrings(XMLDocument *doc);

        void parseWorksheet(XMLDocument *doc);

        void parseWorkbook(XMLDocument *doc);

        void parseContentTypes(XMLDocument *xmlDocument);

        void parseRelationships(XMLDocument *doc);

        void parseRelationShip(XMLDocument *doc, relations_t &relations);

        void parseStyles(XMLDocument *doc);

        void addStyle(XMLElement *element);

        void setDefaultSettings(XMLElement *element);

        void parseAppFile(XMLDocument *doc);

        void parseCoreFile(XMLDocument *doc);

        static void parseWordApp(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

        static void parseWordCore(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

        static void parsePptApp(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

        static void parsePptCore(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

        static void parseXlsxApp(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

        static void parseXlsxCore(tinyxml2::XMLDocument &doc, fileMetaData_t &data);

    public:
        explicit DocumentLoader(options_t &options);

        void loadData();

        XMLDocument *getMainDocument();

        docInfo_t getDocxData() const;

        xlsInfo_t getXlsxData() const;

        pptInfo_t getPptxData() const;
    };

}
#endif //DOCXTOTXT_DOCUMENTLOADER_H
