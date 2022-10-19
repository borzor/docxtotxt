//
// Created by borzor on 10/18/22.
//

#ifndef DOCXTOTXT_TABLECOMMONS_H
#define DOCXTOTXT_TABLECOMMONS_H

enum tableProperty {
    tblJc,
    tblJshd,
    tblBorders,
    tblCaption,
    tblCellMar,
    tblCellSpacing,
    tblInd,
    tblLayout,
    tblLook,
    tblOverlap,
    tblpPr,
    tblStyle,
    tblStyleColBandSize,
    tblStyleRowBandSize,
    tblW
};

enum class tableJustify {
    start,
    end,
    center
};

typedef struct {
    size_t ind;
} tableIndent;

typedef struct {

} floatingSettings;

typedef struct {
    tableJustify justify;
    tableIndent ind;
    floatingSettings floatTable;
    //maybe table width
} tableProperties;


#endif //DOCXTOTXT_TABLECOMMONS_H