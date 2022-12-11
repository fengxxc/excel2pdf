package com.github.fengxxc.util;

/**
 * @author fengxxc
 * @date 2022-12-09
 */
public class Position {
    private int rowIndex;
    private int colIndex;

    public Position() {
    }

    public Position(int colIndex, int rowIndex) {
        this.colIndex = colIndex;
        this.rowIndex = rowIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public Position setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
        return this;
    }

    public int getColIndex() {
        return colIndex;
    }

    public Position setColIndex(int colIndex) {
        this.colIndex = colIndex;
        return this;
    }

    @Override
    public String toString() {
        return "Position{" +
                "y=" + rowIndex +
                ", x=" + colIndex +
                '}';
    }
}
