cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 20)

set(CMAKE_MODULE_PATH "/usr/share/cmake-3.16/Modules" ${CMAKE_MODULE_PATH})
find_package(TinyXML2 REQUIRED)
find_package(LibZip REQUIRED)

add_executable(docxtotxt main.cpp mainDoc/source/MainDocParser.cpp mainDoc/headers/MainDocParser.h mainDoc/source/ParagraphParser.cpp
        mainDoc/headers/ParagraphParser.h mainDoc/source/SectionParser.cpp mainDoc/headers/SectionParser.h mainDoc/source/DrawingParser.cpp
        mainDoc/headers/DrawingParser.h mainDoc/source/TableParser.cpp mainDoc/headers/TableParser.h)
target_link_libraries(docxtotxt PUBLIC tinyxml2::tinyxml2 zip)
