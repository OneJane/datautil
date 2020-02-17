package com.onejane.data;

import com.onejane.head.TheadColumnTree;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DataColumnBlock {

    private TheadColumnTree theadColumnTree ;

    private int startRowIndex;

    private int endRowIndex;

    public DataColumnBlock() {
    }

    public DataColumnBlock(TheadColumnTree theadColumnTree , int startRowIndex , int endRowIndex  ) {
        this.theadColumnTree=theadColumnTree;
        this.startRowIndex=startRowIndex;
        this.endRowIndex=endRowIndex;
    }

}
