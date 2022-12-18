//
// Created by boris on 2/9/22.
//

#ifndef DOCXTOTXT_MAINDOCPARSER_H
#define DOCXTOTXT_MAINDOCPARSER_H


#include "../ParserCommons/DocumentCommons.h"

namespace docxtotxt {
    using namespace std;
    using namespace tinyxml2;

    /*!
	\brief Класс для вызова загрузчка и обработки сторонних файлов
    */
    class MainDocParser {
    private:
        options_t &options;
        wstringConvert convertor;

        void insertHyperlinks(std::map<std::string, std::string> hyperlinkRelationship);

        void saveImages(const std::map<std::string, std::string> &imageRelationship) const;

    public:
        explicit MainDocParser(options_t &options);

        /*!
         *Вызывает загрузчик обрабатывающий входные файлы и выводящий результат в поток вывода
         */
        void parseFile();

    };


}

#endif //DOCXTOTXT_MAINDOCPARSER_H
