//
// Created by borzor on 12/16/22.
//
#include <cmath>
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
/*!
 * Функция для установки настроек выравнивания
 * @param jc Элемент, содержащий настройки выравнивания
 * @param settings Структура настроек
 */
static void setJustify(tinyxml2::XMLElement *jc, docxtotxt::justify_t &settings) {
    if (jc != nullptr) {
        std::string justify = jc->Attribute("w:val");
        if (!strcmp(justify.c_str(), "left"))
            settings = docxtotxt::justify_t::left;
        else if (!strcmp(justify.c_str(), "right"))
            settings = docxtotxt::justify_t::right;
        else if (!strcmp(justify.c_str(), "center"))
            settings = docxtotxt::justify_t::center;
        else if (!strcmp(justify.c_str(), "both"))
            settings = docxtotxt::justify_t::left;
    }
}
/*!
 * Функция для установки настроек отступа
 * @param jc Элемент, содержащий настройки отступа
 * @param settings Структура настроек
 */
static void setIndentation(tinyxml2::XMLElement *ind, docxtotxt::indentation &settings) {
    if (ind != nullptr) {
        auto left = ind->Attribute("w:left");
        if (left != nullptr) {
            settings.left = atoi(left) / TWIP_TO_CHARACTER;
        }
        auto right = ind->Attribute("w:right");
//        if (right != nullptr) {
//            settings.right = atoi(right) / TWIP_TO_CHARACTER;
//        }
        auto hanging = ind->Attribute("w:hanging");
        if (hanging != nullptr) {
            settings.hanging = atoi(hanging) / TWIP_TO_CHARACTER;
        }
        auto firstLine = ind->Attribute("w:firstLine");
        if (firstLine != nullptr) {
            settings.firstLine = atoi(firstLine) / TWIP_TO_CHARACTER;
        }
    }
}
/*!
 * Функция для установки настроек межстрочного отступа
 * @param jc Элемент, содержащий настройки омежстрочного тступа
 * @param settings Структура настроек
 */
static void setSpacing(tinyxml2::XMLElement *spacing, docxtotxt::lineSpacing &settings) {
    if (spacing != nullptr) {
        auto before = spacing->Attribute("w:before");
        if (before != nullptr) {
            settings.before = std::lround(atoi(before) / TWIP_TO_CHARACTER);
        }
        auto after = spacing->Attribute("w:after");
        if (after != nullptr) {
            settings.after = std::lround(atoi(after) / TWIP_TO_CHARACTER);
        }
    }
}
/*!
 * Функция для установки настроек заполнения
 * @param jc Элемент, содержащий настройки заполнения
 * @param settings Структура настроек
 */
static void setTabulation(tinyxml2::XMLElement *tabs, std::vector<docxtotxt::tabulation> &settings) {
    if (tabs != nullptr) {
        auto *tab = tabs->FirstChildElement();
        while (tab != nullptr) {
            docxtotxt::tabulation tmpTab;
            tmpTab.character = docxtotxt::none;
            auto leader = tab->Attribute("w:leader");
            if (leader != nullptr) {
                if (!strcmp(leader, "dot")) {
                    tmpTab.character = docxtotxt::dot;
//                } else if (!strcmp(leader, "heavy")) {
//                    tmpTab.character = docxtotxt::heavy;
//                } else if (!strcmp(leader, "hyphen")) {
//                    tmpTab.character = docxtotxt::hyphen;
//                } else if (!strcmp(leader, "middleDot")) {
//                    tmpTab.character = docxtotxt::middleDot;
//                } else if (!strcmp(leader, "none")) {
//                    tmpTab.character = docxtotxt::none;
//                } else if (!strcmp(leader, "underScore")) {
//                    tmpTab.character = docxtotxt::underscore;
                }
            }
            auto pos = tab->Attribute("w:pos");
            if (pos != nullptr) {
                tmpTab.pos = atoi(pos) / TWIP_TO_CHARACTER;
            }
            settings.emplace_back(tmpTab);
            tab = tab->NextSiblingElement();
        }
    }
}


#endif //DOCXTOTXT_COMMONFUNCTIONS_H
