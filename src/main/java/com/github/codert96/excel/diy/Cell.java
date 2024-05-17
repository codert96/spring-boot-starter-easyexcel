package com.github.codert96.excel.diy;

import lombok.*;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.Date;
import java.util.Objects;
import java.util.function.Function;

@Getter
@ToString
@Accessors(chain = true)
@RequiredArgsConstructor
@EqualsAndHashCode(exclude = "value")
public class Cell implements Comparable<Cell>, Comparator<Cell>, Serializable {
    private final int sheetIndex;
    private final String sheetName;
    private final int rowIndex;
    private final int cellIndex;
    @Setter
    private String value;
    @Setter
    private String link;
    @Setter
    private String comment;
    @Setter
    private String formula;

    public BigDecimal getNumberValue() {
        return new BigDecimal(value);
    }

    public boolean isNumberValue() {
        if (value == null || value.trim().isEmpty()) {
            return false;
        }
        return value.matches("\\d+(\\.\\d+)?");
    }

    public <T> T getValue(Function<String, T> function) {
        return function.apply(value);
    }

    public Date getDateValue(String patten) throws ParseException {
        return new SimpleDateFormat(patten).parse(value);
    }

    public LocalDateTime getLocalDateTimeValue(String patten) {
        return LocalDateTime.from(DateTimeFormatter.ofPattern(patten).parse(value));
    }

    public boolean getBooleanValue() {
        return getValue(Boolean::parseBoolean);
    }

    public boolean getBooleanValue(Function<String, Boolean> function) {
        return getValue(function);
    }

    @Override
    public int compareTo(Cell other) {
        if (!Objects.equals(sheetIndex, other.sheetIndex)) {
            return Integer.compare(sheetIndex, other.sheetIndex);
        }

        if (!Objects.equals(rowIndex, other.rowIndex)) {
            return Integer.compare(rowIndex, other.rowIndex);
        }
        return Integer.compare(cellIndex, other.cellIndex);
    }

    @Override
    public int compare(Cell o1, Cell o2) {
        return o1.compareTo(o2);
    }
}