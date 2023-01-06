cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
#set(CMAKE_CXX_FLAGS_DEBUG  "-O1")

add_subdirectory(Libs/tinyxml2)

add_executable(docxtotxt main.cpp mainDoc/source/MainDocParser.cpp mainDoc/headers/MainDocParser.h
        mainDoc/ParserCommons/DocumentCommons.h mainDoc/ParserCommons/DocumentMetaDataCommons.h
        mainDoc/source/XlsParser.cpp mainDoc/headers/XlsParser.h mainDoc/source/DocumentLoader.cpp mainDoc/headers/DocumentLoader.h
        mainDoc/source/PptParser.cpp mainDoc/headers/PptParser.h mainDoc/source/DocParser.cpp mainDoc/headers/DocParser.h
        mainDoc/ParserCommons/FileSpecificCommons/PptCommons.h mainDoc/ParserCommons/FileSpecificCommons/XlsCommons.h
        mainDoc/ParserCommons/FileSpecificCommons/DocCommons.h mainDoc/ParserCommons/CommonFunctions.h
        mainDoc/source/BufferWriter.cpp mainDoc/ParserCommons/BufferWriter.h mainDoc/headers/Parser.h)

#target_link_libraries(docxtotxt PUBLIC libzip tinyxml2)
target_link_libraries(docxtotxt ${CMAKE_SOURCE_DIR}/Libs/libzip.so ${CMAKE_SOURCE_DIR}/Libs/libtinyxml2.a )

#target_link_libraries(docxtotxt PUBLIC tinyxml2)
#target_link_libraries(docxtotxt PUBLIC libzip)
#target_link_libraries(docxtotxt PUBLIC libzip tinyxml2) # ${CMAKE_SOURCE_DIR}/Libs/tinyxml2/cmake-build-debug/libtinyxml2.a

