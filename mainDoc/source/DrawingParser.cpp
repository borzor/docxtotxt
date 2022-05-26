//
// Created by boris on 2/22/22.
//

#include "../headers/DrawingParser.h"

namespace Drawing {
    DrawingParser::DrawingParser() {
    }

    void DrawingParser::parseDrawing(tinyxml2::XMLElement *element) {
        auto current_element = element->FirstChildElement()->FirstChildElement();

        while (current_element != nullptr) {

            if (!strcmp(current_element->Value(), "a:graphic")) {
                if (current_element->FirstChildElement("a:graphicData") != nullptr) {
                    if (current_element->FirstChildElement("a:graphicData")->FirstChildElement("pic:pic") != nullptr) {
                        auto picElement = current_element->FirstChildElement("a:graphicData")->FirstChildElement(
                                "pic:pic")->FirstChildElement();
                        while (picElement != nullptr) {
                            if (!strcmp(picElement->Value(), "pic:nvPicPr")) {
                                //maybe  Specifies non-visual properties of a picture, such as locking properties, name, id and title, and whether the picture is hidden.
                            } else if (!strcmp(picElement->Value(), "pic:blipFill")) {
                                lastImageId = picElement->FirstChildElement()->Attribute("r:embed");
                                //TODO pointer to picture
                            } else if (!strcmp(picElement->Value(), "pic:spPr")) {
                                if (picElement->FirstChildElement("a:xfrm") != nullptr) {
                                    if (picElement->FirstChildElement("a:xfrm")->FirstChildElement("a:ext") !=
                                        nullptr) {
                                        this->height = atoi(picElement->FirstChildElement("a:xfrm")->FirstChildElement(
                                                "a:ext")->Attribute("cy"));
                                        this->width = atoi(picElement->FirstChildElement("a:xfrm")->FirstChildElement(
                                                "a:ext")->Attribute("cx"));
                                    }
                                }
                            }
                            picElement = picElement->NextSiblingElement();
                        }
                    }
                }
            }
            current_element = current_element->NextSiblingElement();
        }
    }

    void DrawingParser::getDrawingSize(size_t &height, size_t &width) const {
        height = this->height;
        width = this->width;
    }

    string DrawingParser::getDrawingId() const {
        return this->lastImageId;
    }


}

