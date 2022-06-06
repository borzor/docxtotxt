//
// Created by borzor on 12/16/22.
//
#include "DocumentCommons.h"
#include "tinyxml2.h"

#ifndef DOCXTOTXT_COMMONFUNCTIONS_H
#define DOCXTOTXT_COMMONFUNCTIONS_H
namespace docxtotxt {
/*!
 * \brief Функция для извлечения информации координат объекта
 * @param xfrm Элемент, дочерние элементы которого содержат информацию о координатах
 * @return
 */
    static objectInfo_t extractObjectInfo(tinyxml2::XMLElement *xfrm) {
        objectInfo_t object;
        if (xfrm != nullptr) {
            auto off = xfrm->FirstChildElement("a:off");
            auto ext = xfrm->FirstChildElement("a:ext");
            if (off != nullptr) {
                object.offsetX = atoi(off->Attribute("x"));
                object.offsetY = atoi(off->Attribute("y"));
            }
            if (ext != nullptr) {
                object.objectSizeX = atoi(ext->Attribute("cx"));
                object.objectSizeY = atoi(ext->Attribute("cy"));
            }
        }
        return object;
    }

/*!
 * \brief Функция для извлечения picture объекта
 * @param element Элемент, дочерние элементы которого содержат информацию о изображении
 * @param prefixTag Префикс тэга, зависящий от исходного документа
 * @return
 */
    static picture extractPicture(tinyxml2::XMLElement *element, const std::string &prefixTag) {
        picture object;
        auto spPr = element->FirstChildElement(std::string(prefixTag + ":spPr").c_str());
        auto blipFill = element->FirstChildElement(std::string(prefixTag + ":blipFill").c_str());
        if (blipFill != nullptr) {
            auto blip = blipFill->FirstChildElement("a:blip");
            if (blip != nullptr) {
                object.rId = blip->Attribute("r:embed");
            }
        }
        if (spPr != nullptr) {
            auto xfrm = spPr->FirstChildElement("a:xfrm");
            object.objectInfo = extractObjectInfo(xfrm);
        }
        return object;
    }
}


#endif //DOCXTOTXT_COMMONFUNCTIONS_H
