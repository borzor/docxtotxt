cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
#set(CMAKE_CXX_FLAGS_DEBUG "-g")
set(CMAKE_CXX_FLAGS_DEBUG  "-O3")

add_executable(docxtotxt main.cpp mainDoc/source/MainDocParser.cpp mainDoc/headers/MainDocParser.h
        mainDoc/ParserCommons/DocumentCommons.h mainDoc/ParserCommons/DocumentMetaDataCommons.h
        mainDoc/source/XlsParser.cpp mainDoc/headers/XlsParser.h mainDoc/source/DocumentLoader.cpp mainDoc/headers/DocumentLoader.h
        mainDoc/source/PptParser.cpp mainDoc/headers/PptParser.h mainDoc/source/DocParser.cpp mainDoc/headers/DocParser.h
        mainDoc/ParserCommons/FileSpecificCommons/PptCommons.h mainDoc/ParserCommons/FileSpecificCommons/XlsCommons.h
        mainDoc/ParserCommons/FileSpecificCommons/DocCommons.h mainDoc/ParserCommons/CommonFunctions.h
        mainDoc/source/BufferWriter.cpp mainDoc/ParserCommons/BufferWriter.h mainDoc/headers/Parser.h )#Libs/tinyxml2.h Libs/tinyxml2.cpp

target_link_libraries(docxtotxt ${CMAKE_SOURCE_DIR}/Libs/libtinyxml2.so  ${CMAKE_SOURCE_DIR}/Libs/libzip.so)

