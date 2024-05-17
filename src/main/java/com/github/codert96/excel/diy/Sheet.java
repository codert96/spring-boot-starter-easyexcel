package com.github.codert96.excel.diy;

import lombok.*;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.util.*;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Getter
@AllArgsConstructor
@EqualsAndHashCode(callSuper = false)
public class Sheet extends HashMap<Integer, Row> implements Comparable<Sheet>, Comparator<Sheet>, Serializable {
    private final int sheetIndex;
    private final String sheetName;

    @Getter(AccessLevel.NONE)
    private final Set<Merge> merges = new HashSet<>();

    public Merge put(Merge merge) {
        merges.add(merge);
        return merge;
    }

    public String getMergeValue(int rowindex, int colindex) {
        return stream(merges)
                .filter(m -> m.isMergeRange(rowindex, colindex))
                .findFirst()
                .map(Merge::getValue)
                .orElse(null);
    }

    public boolean isMergeRange(int rowindex, int colindex) {
        return stream(merges).anyMatch(m -> m.isMergeRange(rowindex, colindex));
    }

    public Row getRow(int rowindex) {
        return get(rowindex);
    }

    public Row findRow(Predicate<Row> consumer) {
        return stream(values()).filter(consumer).sorted().findFirst().orElse(null);
    }

    public List<Row> findRows(Predicate<Row> consumer) {
        return stream(values()).filter(consumer).sorted().collect(Collectors.toCollection(LinkedList::new));
    }


    private <T> Stream<T> stream(Collection<T> collection) {
        int size = collection.size();
        return (size > 1024 ? collection.parallelStream() : collection.stream());
    }

    @Override
    public int compareTo(Sheet o) {
        return Integer.compare(sheetIndex, o.sheetIndex);
    }

    @Override
    public int compare(Sheet o1, Sheet o2) {
        return o1.compareTo(o2);
    }


    @Getter
    @Accessors(chain = true)
    @EqualsAndHashCode(exclude = "value")
    @RequiredArgsConstructor
    public static class Merge implements Serializable {
        private final int firstRowIndex;

        private final int lastRowIndex;

        private final int firstColumnIndex;

        private final int lastColumnIndex;

        @Setter
        private String value;
        @Setter
        private String link;
        @Setter
        private String comment;
        @Setter
        private String formula;

        public boolean isMergeRange(int rowindex, int colindex) {
            return rowindex >= firstRowIndex && rowindex <= lastRowIndex &&
                    colindex >= firstColumnIndex && colindex <= lastColumnIndex;
        }
    }

}