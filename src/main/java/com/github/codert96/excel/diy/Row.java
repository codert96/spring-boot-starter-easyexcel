package com.github.codert96.excel.diy;

import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;

import java.io.Serializable;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.function.Predicate;
import java.util.stream.Collectors;

@Getter
@AllArgsConstructor
@EqualsAndHashCode(callSuper = false)
public class Row extends HashMap<Integer, Cell> implements Comparable<Row>, Comparator<Row>, Serializable {
    private final int sheetIndex;
    private final String sheetName;
    private final int rowIndex;


    public Cell findCell(Predicate<Cell> predicate) {
        return values().stream().filter(predicate).findFirst().orElse(null);
    }

    public Cell getCell(int colindex) {
        return get(colindex);
    }

    public List<Cell> findCells(Predicate<Cell> predicate) {
        return values().stream().filter(predicate).sorted().collect(Collectors.toCollection(LinkedList::new));
    }

    @Override
    public int compareTo(Row other) {
        if (sheetIndex != other.sheetIndex) {
            return Integer.compare(sheetIndex, other.sheetIndex);
        } else {
            return Integer.compare(rowIndex, other.rowIndex);
        }
    }

    @Override
    public int compare(Row o1, Row o2) {
        return o1.compareTo(o2);
    }

}