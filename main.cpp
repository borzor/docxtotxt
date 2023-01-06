#include "mainDoc/headers/MainDocParser.h"
#include <getopt.h>
#include <deque>
#include <sys/stat.h>
#include <chrono>
#include <iostream>
#include <fstream>

#define OPTSTR "i:d:o:lmhnr"
using namespace std;

//void coverageTest(docxtotxt::options_t options) {
//    auto start = std::chrono::steady_clock::now();
//    docxtotxt::MainDocParser parser(options);
//    parser.parseFile();
//    std::cout << "Elapsed(ms)="
//              << std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count()
//              << std::endl;
//
//}
//
//docxtotxt::options_t
//prepareOptions(const char *fileName, const char *output = nullptr, const char *imageDir = nullptr) {
//    docxtotxt::options_t options = {&std::wcout};
//    if (docxtotxt::ends_with(fileName, ".docx")) {
//        options.docType = docxtotxt::docx;
//    } else if (docxtotxt::ends_with(fileName, ".pptx")) {
//        options.docType = docxtotxt::pptx;
//    } else if (docxtotxt::ends_with(fileName, ".xlsx")) {
//        options.docType = docxtotxt::xlsx;
//    }
//    options.filePath = fileName;
//    options.flags |= 1 << 0;
//    options.flags |= 1 << 2;
//    options.flags |= 1 << 3;
//    options.flags |= 1 << 4;
//    options.flags |= 1 << 5;
//    if (output != nullptr) {
//        std::locale loco(std::locale(), new std::codecvt_utf8<wchar_t>);
//        options.output = new std::wofstream(output);
//        options.output->imbue(loco);
//    }
//    if (imageDir != nullptr) {
//        options.flags |= 1 << 1;
//        mkdir(imageDir, 0777);
//        options.pathToDraws = std::string(imageDir);
//    }
//    return options;
//}
//
//void parseFullPresentation() {
//    auto options = prepareOptions("/home/borzor/Documents/pptx/presentation_2.pptx",
//                                  "/home/borzor/Documents/pptx/presentation_2/pptx.txt",
//                                  "/home/borzor/Documents/pptx/presentation_2");
//    coverageTest(options);
//}
//
//void parseFullPresentation2() {
//    auto options = prepareOptions("/home/borzor/Documents/pptx/presentation_1.pptx",
//                                  "/home/borzor/Documents/pptx/presentation_1/pptx.txt",
//                                  "/home/borzor/Documents/pptx/presentation_1");
//    coverageTest(options);
//}
//
//void parseFullDocx() {
//    auto options = prepareOptions("/home/borzor/Documents/docx/docx_1.docx",
//                                  "/home/borzor/Documents/docx/docx_1/docx.txt", "/home/borzor/Documents/docx/docx_1");
//    options.flags ^= 1UL << 5;
//    coverageTest(options);
//}
//
//void parseFullDocx2() {
//    auto options = prepareOptions("/home/borzor/Documents/docx/docx_2.docx",
//                                  "/home/borzor/Documents/docx/docx_2/docx.txt", "/home/borzor/Documents/docx/docx_2");
//    coverageTest(options);
//}
//
//void parseFullDocx3() {
//    auto options = prepareOptions("/home/borzor/Documents/docx/docx_3.docx",
//                                  "/home/borzor/Documents/docx/docx_3/docx.txt", "/home/borzor/Documents/docx/docx_3");
//    options.flags ^= 1UL << 5;
//    coverageTest(options);
//}
//
//void parseFullXlsx() {
//    auto options = prepareOptions("/home/borzor/Documents/xlsx/excel_4.xlsx",
//                                  "/home/borzor/Documents/xlsx/excel_4/excel.txt",
//                                  "/home/borzor/Documents/xlsx/excel_4");
//    coverageTest(options);
//}
//
//void parseFullXlsx2() {
//    auto options = prepareOptions("/home/borzor/Documents/xlsx/excel_1.xlsx",
//                                  "/home/borzor/Documents/xlsx/excel_1/excel.txt",
//                                  "/home/borzor/Documents/xlsx/excel_1");
//    options.flags ^= 1UL << 5;
//    coverageTest(options);
//}
//
//
//int main() {
//    parseFullDocx();
//    parseFullDocx2();
//    parseFullDocx3();
//    parseFullPresentation();
//    parseFullPresentation2();
//    parseFullXlsx();
//    parseFullXlsx2();
//    return 0;
//}

int main(int argc, char *argv[]) {
    auto start = std::chrono::steady_clock::now();
    int opt;
    docxtotxt::options_t options = {&std::wcout};
    while ((opt = getopt(argc, argv, OPTSTR)) != EOF)
        switch (opt) {
            case 'i': {//input file
                auto fileName = std::string(optarg);
                if (docxtotxt::ends_with(fileName, ".docx")) {
                    options.docType = docxtotxt::docx;
                } else if (docxtotxt::ends_with(fileName, ".pptx")) {
                    options.docType = docxtotxt::pptx;
                } else if (docxtotxt::ends_with(fileName, ".xlsx")) {
                    options.docType = docxtotxt::xlsx;
                } else {
                    throw std::invalid_argument("No input file by parameter -i");
                }
                options.filePath = optarg;
                break;
            }
            case 'd': {//draws
                options.flags |= 1 << 1;
                mkdir(optarg, 0777);
                options.pathToDraws = std::string(optarg);
                break;
            }
            case 'o': {//output file
                std::locale loco(std::locale(), new std::codecvt_utf8<wchar_t>);
                options.flags |= 1 << 0;
                options.output = new std::wofstream(optarg);
                options.output->imbue(loco);
                break;
            }
            case 'l': {//external hyperlinks
                options.flags |= 1 << 2;
                break;
            }
            case 'n': {//presentation notes
                options.flags |= 1 << 3;
                break;
            }
            case 'm': {//doc metadata
                options.flags |= 1 << 4;
                break;
            }
            case 'r': {//raw cells data in tables in xlsx
                options.flags |= 1 << 5;
                break;
            }
            case 'h': {
                std::string usage = "usage: docxtotxt ";
                std::cout << usage << "[-lhmnr] [-d image_dir] [-i input_file]" << std::endl
                          << std::string(usage.length(), ' ') << "[-o output_file]" << std::endl;
                return 0;
//                if (options.output != nullptr)
            }
            default:
                std::cerr << "Invalid parameters" << std::endl;
        }
    try {
        if (options.filePath == nullptr) {
            throw std::invalid_argument("No input file by parameter -i");
        }
        docxtotxt::MainDocParser parser(options);
        parser.parseFile();
        std::cout << "Elapsed(ms)="
                  << std::chrono::duration_cast<std::chrono::milliseconds>(
                          std::chrono::steady_clock::now() - start).count()
                  << std::endl;
    } catch (std::exception &e) {
        std::cout << e.what() << std::endl;
    }
    return 0;
}
