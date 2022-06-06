cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 11)

#set(CMAKE_MODULE_PATH "/usr/share/cmake/modules" ${CMAKE_MODULE_PATH})
#find_package(TinyXML2 REQUIRED)
#find_package(LibZip REQUIRED)

add_executable(docxtotxt main.cpp mainDoc/source/MainDocParser.cpp mainDoc/headers/MainDocParser.h mainDoc/source/ParagraphParser.cpp
        mainDoc/headers/ParagraphParser.h
        mainDoc/source/TableParser.cpp mainDoc/headers/TableParser.h mainDoc/ParserCommons/ParagraphCommons.h
        mainDoc/ParserCommons/DocumentCommons.h mainDoc/ParserCommons/TableCommons.h mainDoc/ParserCommons/DocumentMetaDataCommons.h
        mainDoc/source/XlsParser.cpp mainDoc/headers/XlsParser.h mainDoc/source/DocumentLoader.cpp mainDoc/headers/DocumentLoader.h
        mainDoc/source/PptParser.cpp mainDoc/headers/PptParser.h mainDoc/source/DocParser.cpp mainDoc/headers/DocParser.h
        mainDoc/ParserCommons/FileSpecificCommons/PptCommons.h mainDoc/ParserCommons/FileSpecificCommons/XlsCommons.h
        mainDoc/ParserCommons/FileSpecificCommons/DocCommons.h mainDoc/ParserCommons/CommonFunctions.h
        mainDoc/source/BufferWriter.cpp mainDoc/ParserCommons/BufferWriter.h mainDoc/headers/Parser.h)
#target_link_libraries(docxtotxt PUBLIC tinyxml2::tinyxml2 zip)
message(${CMAKE_SOURCE_DIR})
target_link_libraries(docxtotxt ${CMAKE_SOURCE_DIR}/Libs/libtinyxml2.so ${CMAKE_SOURCE_DIR}/Libs/libzip.so)