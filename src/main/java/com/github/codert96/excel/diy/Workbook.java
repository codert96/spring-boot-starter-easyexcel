package com.github.codert96.excel.diy;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.FormulaData;
import com.alibaba.excel.read.metadata.holder.ReadSheetHolder;
import lombok.EqualsAndHashCode;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.io.Serializable;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Predicate;
import java.util.stream.Collectors;

@Slf4j
@EqualsAndHashCode(callSuper = false)
public class Workbook extends AnalysisEventListener<Map<Integer, String>> implements Map<Integer, Sheet>, Serializable {
    private final Map<Integer, Sheet> excel = new HashMap<>();

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {

        ReadSheetHolder currentReadHolder = context.readSheetHolder();

        String sheetName = currentReadHolder.getSheetName();
        int sheetNo = currentReadHolder.getSheetNo();

        context.readRowHolder().getCellMap().forEach((k, cell) -> {
            CellData<?> cellData = (CellData<?>) cell;
            put(sheetNo, sheetName, cell.getRowIndex(), cell.getColumnIndex(), cellData.getStringValue())
                    .setFormula(
                            Optional.of(cellData).map(CellData::getFormulaData)
                                    .map(FormulaData::getFormulaValue).orElse(null)
                    )
            ;
        });
    }

    private Cell put(int sheetIndex, String sheetName, int rowindex, int colindex, String value) {
        return excel.computeIfAbsent(sheetIndex, it -> new Sheet(it, sheetName))
                .computeIfAbsent(rowindex, it -> new Row(sheetIndex, sheetName, it))
                .computeIfAbsent(colindex, it -> new Cell(sheetIndex, sheetName, rowindex, colindex))
                .setValue(value);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        ReadSheetHolder currentReadHolder = (ReadSheetHolder) context.currentReadHolder();

        Sheet sheet = getSheetAt(currentReadHolder.getSheetNo());

        if (CellExtraTypeEnum.MERGE.equals(extra.getType())) {
            Sheet.Merge merge = sheet.put(
                    new Sheet.Merge(
                            extra.getFirstRowIndex(),
                            extra.getLastRowIndex(),
                            extra.getFirstColumnIndex(),
                            extra.getLastColumnIndex()
                    )
            );
            CellExtraEach rangeForeach = new CellExtraEach(sheet, extra);

            AtomicReference<Cell> atomicReference = new AtomicReference<>();

            rangeForeach.accept((rowindex, colindex, cell) -> {
                if (cell != null && cell.getValue() != null) {
                    atomicReference.set(cell);
                }
            });

            if (Objects.nonNull(atomicReference.get())) {
                Cell ref = atomicReference.get();
                merge.setValue(ref.getValue())
                        .setLink(ref.getLink())
                        .setComment(ref.getComment())
                        .setFormula(ref.getFormula());
                rangeForeach.accept((rowindex, colindex, cell) -> {
                    if (cell == null) {
                        put(ref.getSheetIndex(), ref.getSheetName(), rowindex, colindex, ref.getValue())
                                .setLink(ref.getLink())
                                .setComment(ref.getComment())
                                .setFormula(ref.getFormula());
                    } else if (cell.getValue() == null) {
                        cell.setValue(ref.getValue())
                                .setLink(ref.getLink())
                                .setComment(ref.getComment())
                                .setFormula(ref.getFormula());
                    }
                });
            }
        } else if (CellExtraTypeEnum.HYPERLINK.equals(extra.getType())) {
            CellExtraEach rangeForeach = new CellExtraEach(sheet, extra);
            rangeForeach.accept(
                    (rowindex, colindex, cell) -> {
                        cell.setLink(extra.getText());
                    }
            );
        } else if (CellExtraTypeEnum.COMMENT.equals(extra.getType())) {
            CellExtraEach rangeForeach = new CellExtraEach(sheet, extra);
            rangeForeach.accept(
                    (rowindex, colindex, cell) -> {
                        cell.setLink(extra.getText());
                    }
            );
        }
    }

    public Collection<Sheet> getSheets() {
        return excel.values();
    }

    public Sheet getSheetAt(int index) {
        return get(index);
    }

    public Sheet getSheet(String name) {
        return getSheet(sheet -> name.equals(sheet.getSheetName()));
    }

    public Sheet getSheet(Predicate<Sheet> predicate) {
        return excel.values().stream().filter(predicate).findFirst().orElse(null);
    }

    public List<Sheet> getSheets(Predicate<Sheet> predicate) {
        return excel.values().stream().filter(predicate).collect(Collectors.toCollection(LinkedList::new));
    }

    public List<String> getSheetNames() {
        return excel.values()
                .stream()
                .map(Sheet::getSheetName)
                .collect(Collectors.toCollection(LinkedList::new));
    }

    @Override
    public int size() {
        return excel.size();
    }

    @Override
    public boolean isEmpty() {
        return excel.isEmpty();
    }

    @Override
    public boolean containsKey(Object key) {
        return excel.containsKey(key);
    }

    @Override
    public boolean containsValue(Object value) {
        return excel.containsValue(value);
    }

    @Override
    public Sheet get(Object key) {
        return excel.get(key);
    }

    @Override
    public Sheet put(Integer key, Sheet value) {
        return excel.put(key, value);
    }

    @Override
    public Sheet remove(Object key) {
        return excel.remove(key);
    }

    @Override
    public void putAll(Map<? extends Integer, ? extends Sheet> m) {
        excel.putAll(m);
    }

    @Override
    public void clear() {
        excel.clear();
    }

    @Override
    public Set<Integer> keySet() {
        return excel.keySet();
    }

    @Override
    public Collection<Sheet> values() {
        return excel.values();
    }

    @Override
    public Set<Entry<Integer, Sheet>> entrySet() {
        return excel.entrySet();
    }

    @RequiredArgsConstructor
    private static class CellExtraEach {
        private final Sheet sheet;
        private final CellExtra extra;

        public void accept(EachCell eachCell) {
            for (int i = extra.getFirstRowIndex(); i <= extra.getLastRowIndex(); i++) {
                for (int j = extra.getFirstColumnIndex(); j <= extra.getLastColumnIndex(); j++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        Cell cell = row.getCell(j);
                        eachCell.accept(i, j, cell);
                    } else {
                        eachCell.accept(i, j, null);
                    }
                }
            }
        }
    }

    private interface EachCell {
        void accept(int rowindex, int colindex, Cell cell);
    }
}