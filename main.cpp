#include "mainDoc/headers/MainDocParser.h"
#include <getopt.h>
#include <deque>
#include <sys/stat.h>
#include <chrono>

#define OPTSTR "i:d:o:h"


int main(int argc, char *argv[]) {
    auto start = std::chrono::steady_clock::now();
    int opt;
    options_t options = {0x0};
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
                options.flags |= 1 << 0;
                options.output = std::ofstream(optarg);
                break;
            }
            case 'h':
                //TODO help options...
            default:
                std::cerr << "Invalid parameters" << std::endl;

        }
    prsr::MainDocParser parser(std::move(options));
    try {
        parser.doInit();
        parser.parseMainDoc();
    } catch (std::exception &e) {
        std::cout << e.what() << std::endl;
    }
    std::cout << "Elapsed(ms)="
              << std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count()
              << std::endl;


}