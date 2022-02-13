#include "parser.h"

int main(int argc, char *argv[]) {
    prsr::parser parser("testParser");
    try {
        parser.doInit(argv[1]);
        parser.parseMainDoc();
    } catch (std::exception &e) {
        std::cout << e.what()<<std::endl;
    }

}