cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(docxtotxt main.cpp mainDoc/source/MainDocParser.cpp mainDoc/headers/MainDocParser.h
        mainDoc/ParserCommons/DocumentCommons.h mainDoc/ParserCommons/DocumentMetaDataCommons.h
        mainDoc/source/XlsParser.cpp mainDoc/headers/XlsParser.h mainDoc/source/DocumentLoader.cpp mainDoc/headers/DocumentLoader.h
        mainDoc/source/PptParser.cpp mainDoc/headers/PptParser.h mainDoc/source/DocParser.cpp mainDoc/headers/DocParser.h
        mainDoc/ParserCommons/FileSpecificCommons/PptCommons.h mainDoc/ParserCommons/FileSpecificCommons/XlsCommons.h
        mainDoc/ParserCommons/FileSpecificCommons/DocCommons.h mainDoc/ParserCommons/CommonFunctions.h
        mainDoc/source/BufferWriter.cpp mainDoc/ParserCommons/BufferWriter.h mainDoc/headers/Parser.h)

target_link_libraries(docxtotxt ${CMAKE_SOURCE_DIR}/Libs/libtinyxml2.so ${CMAKE_SOURCE_DIR}/Libs/libzip.so)

#if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
#    add_compile_options(-fprofile-instr-generate -fcoverage-mapping)
#    add_link_options(-fprofile-instr-generate)
#    #Uncomment in case of linker errors
#    link_libraries(clang_rt.profile-x86_64)
#elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
#    add_compile_options(--coverage)
#    #Uncomment in case of linker errors
#    link_libraries(gcov)
#endif ()