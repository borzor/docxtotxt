#include "mainDoc/headers/MainDocParser.h"
#include <getopt.h>
#include <deque>
#include <sys/stat.h>
#include <chrono>
#include <iostream>
#include <fstream>

#define OPTSTR "i:d:o:lmhn"


int main(int argc, char *argv[]) {
    auto start = std::chrono::steady_clock::now();
    int opt;
    docxtotxt::options_t options = {&std::wcout};
    options.printDocProps = false;
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
                mkdir(optarg, 0777);//TODO change directory rule
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
            case 'm': {
                options.printDocProps = true;
                break;
            }
            case 'h': {
                std::string usage = "usage: docxtotxt ";
                std::cout << usage << "[-lmn] [-d image_dir] [-h help] [-i input_file]" << std::endl
                          << std::string(usage.length(), ' ') << "[-o output_file]" << std::endl;
                return 0;
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
