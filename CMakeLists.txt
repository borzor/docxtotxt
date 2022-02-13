cmake_minimum_required(VERSION 3.21)
project(docxtotxt)

set(CMAKE_CXX_STANDARD 20)

set(CMAKE_MODULE_PATH "/usr/share/cmake-3.16/Modules" ${CMAKE_MODULE_PATH})
find_package(TinyXML2 REQUIRED)
find_package(LibZip REQUIRED)

add_executable(docxtotxt main.cpp parser.cpp parser.h)
target_link_libraries(docxtotxt PUBLIC tinyxml2::tinyxml2 zip)
