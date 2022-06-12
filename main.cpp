#include "mainDoc/headers/MainDocParser.h"
#include <getopt.h>
#include <deque>
#include <sys/stat.h>
#include <chrono>
#include <locale>
#define OPTSTR "i:d:o:h:l"


int main(int argc, char *argv[]) {
    auto start = std::chrono::steady_clock::now();
    int opt;
    options_t options = {0x0, nullptr, &std::wcout};
    while ((opt = getopt(argc, argv, OPTSTR)) != EOF)
        switch (opt) {
            case 'i': {//input file
                int err;
                options.input = zip_open(optarg, 0, &err);
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
            case 'h':
                //TODO help options...
            default:
                std::cerr << "Invalid parameters" << std::endl;

        }

    try {
        if (options.input == nullptr) {
            throw std::invalid_argument("No input file by parameter -i");
        }
        docInfo_t docInfo = {0, 0, 0};
        prsr::MainDocParser parser(options, docInfo);
        parser.doInit();
        parser.parseMainDoc();

    } catch (std::exception &e) {
        std::cout << e.what() << std::endl;
    }
    std::cout << "Elapsed(ms)="
              << std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count()
              << std::endl;


}