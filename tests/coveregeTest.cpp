#include "../mainDoc/headers/MainDocParser.h"
#include <deque>
#include <sys/stat.h>
#include <chrono>
#include <iostream>
#include <fstream>

#define OPTSTR "i:d:o:lmhn"

void coverageTest(docxtotxt::options_t options){
    auto start = std::chrono::steady_clock::now();

    docxtotxt::MainDocParser parser(options);
    parser.parseFile();
    std::cout << "Elapsed(ms)="
              << std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count()
              << std::endl;

}

void parseFullPresentation(){
    docxtotxt::options_t options = {&std::wcout};
    options.filePath = "/home/borzor/Documents/pptx/presentation_4.pptx";
    options.docType = docxtotxt::pptx;
    options.flags |= 1 << 1;
    mkdir("/home/borzor/Documents/pptx/presentation_4", 0777);//TODO change directory rule
    options.pathToDraws = std::string("/home/borzor/Documents/pptx/presentation_4");
    options.flags |= 1 << 3;
    options.flags |= 1 << 2;
    std::locale loco(std::locale(), new std::codecvt_utf8<wchar_t>);
    options.flags |= 1 << 0;
    options.output = new std::wofstream("/home/borzor/Documents/pptx/presentation_4/pptx.txt");
    options.output->imbue(loco);
    coverageTest(options);
}

void parseFullDocx(){
    docxtotxt::options_t options = {&std::wcout};
    options.filePath = "/home/borzor/Documents/docx/docx_1.docx";
    options.docType = docxtotxt::docx;
    options.flags |= 1 << 1;
    mkdir("/home/borzor/Documents/docx/docx_1/", 0777);//TODO change directory rule
    options.pathToDraws = std::string("/home/borzor/Documents/docx/docx_1/");
    options.flags |= 1 << 3;
    options.flags |= 1 << 2;
    std::locale loco(std::locale(), new std::codecvt_utf8<wchar_t>);
    options.flags |= 1 << 0;
    options.output = new std::wofstream("/home/borzor/Documents/docx/docx_1/docx.txt");
    options.output->imbue(loco);
    coverageTest(options);
}

void parseFullXlsx(){
    docxtotxt::options_t options = {&std::wcout};
    options.filePath = "/home/borzor/Documents/xlsx/excel_4.xlsx";
    options.docType = docxtotxt::xlsx;
    options.flags |= 1 << 1;
    mkdir("/home/borzor/Documents/xlsx/excel_4", 0777);//TODO change directory rule
    options.pathToDraws = std::string("/home/borzor/Documents/xlsx/excel_4");
    options.flags |= 1 << 3;
    options.flags |= 1 << 2;
    std::locale loco(std::locale(), new std::codecvt_utf8<wchar_t>);
    options.flags |= 1 << 0;
    options.output = new std::wofstream("/home/borzor/Documents/xlsx/excel_4/excel");
    options.output->imbue(loco);
    coverageTest(options);
}

int main() {
    parseFullPresentation();
    parseFullDocx();
    parseFullXlsx();
    return 0;
}
