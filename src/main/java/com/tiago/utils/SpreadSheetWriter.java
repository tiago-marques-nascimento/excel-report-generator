package com.tiago.utils;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class SpreadSheetWriter {

    private static class MatchToStringParser {
        private final String value;
        private final String pipe;
        private final String defaultValue;

        public MatchToStringParser(String match) {
            String content = match.replace("{", "").replace("}", "");
            this.value = extractValue(content);
            this.defaultValue = extractDefaultValue(content);
            this.pipe = extractPipe(content);
        }

        private static String extractValue(String content) {
            if (containsDefaultValue(content)) {
                return splitByDefaultValue(content)[0];
            } else if (containsPipe(content)) {
                return splitByPipe(content)[0];
            }
            return content;
        }
        private static String extractDefaultValue(String content) {
            if (containsDefaultValue(content)) {
                String defaultValue = splitByDefaultValue(content)[1];
                return containsPipe(defaultValue)
                    ? splitByPipe(defaultValue)[0]
                    : defaultValue;
            }
            return null;
        }

        private static String extractPipe(String content) {
            return containsPipe(content) ? splitByPipe(content)[1] : null;
        }


        private static boolean containsDefaultValue(String matchContent) {
            return matchContent.contains("?");
        }

        private static String[] splitByDefaultValue(String content) {
            return content.split("\\?");
        }

        private static boolean containsPipe(String matchContent) {
            return matchContent.contains("|");
        }

        private static String[] splitByPipe(String content) {
            return content.split("\\|");
        }

        public String getValue() {
            return value;
        }

        public Optional<String> getPipe() {
            return Optional.ofNullable(pipe);
        }

        public Optional<String> getDefaultValue() {
            return Optional.ofNullable(defaultValue);
        }
    }

    private static class SpreadSheetAcessor {
        private Integer rows;
        private Integer columns;
        private Integer currentRow;
        private Integer currentColumn;
        private Integer rowsThreshold;
        private Integer expandedRows;
        private Integer expandedColumns;

        public SpreadSheetAcessor(final Integer rows, final Integer columns) {

            this.rows = rows.intValue();
            this.columns = columns.intValue();
            this.currentRow = 0;
            this.currentColumn = 0;
            this.rowsThreshold = -1;
            this.expandedRows = 0;
            this.expandedColumns = 0;
        }

        public SpreadSheetAcessor(
            final Integer rows,
            final Integer columns,
            final Integer currentRow,
            final Integer currentColumn,
            final Integer rowsThreshold
        ) {

            this.rows = rows.intValue();
            this.columns = columns.intValue();
            this.currentRow = currentRow.intValue();
            this.currentColumn = currentColumn.intValue();
            this.rowsThreshold = rowsThreshold.intValue();
            this.expandedRows = 0;
            this.expandedColumns = 0;
        }

        public SpreadSheetAcessor SubSpreadSheetAcessor(
            final Integer rowsThreshold) {
            
            return new SpreadSheetAcessor(
                this.rows,
                this.columns,
                this.currentRow,
                this.currentColumn,
                rowsThreshold);
        }

        public Integer getRows() {
            return this.rows;
        }

        public Integer getColumns() {
            return this.columns;
        }

        public Integer getCurrentRow() {
            return this.currentRow;
        }

        public Integer getCurrentColumn() {
            return this.currentColumn;
        }

        public Integer getExpandedRows() {
            return this.expandedRows;
        }

        public void expandRows(Integer rows) {
            this.rows += rows;

            if(this.rowsThreshold >= 0) {
                this.rowsThreshold += rows;
            }
            this.expandedRows += rows;
        }

        public void expandColumns(Integer columns) {
            this.columns += columns;
            this.expandedColumns += columns;
        }

        public Boolean isDone() {
            return this.currentRow >= this.rows ||
                ((this.rowsThreshold >= 0) && this.currentRow >= this.rowsThreshold);
        }

        public void next() {
            if(!this.isDone() && ++this.currentColumn == this.columns) {
                this.currentColumn = 0;
                this.expandedColumns = 0;
                this.currentRow++;
            }
        }

        public void reset() {
            this.currentColumn = 0;
            this.currentRow = 0;
        }

        public void copyState(SpreadSheetAcessor spreadSheetAcessor) {
            this.rows = spreadSheetAcessor.getRows().intValue();
            this.columns = spreadSheetAcessor.getColumns().intValue();
            this.currentRow = spreadSheetAcessor.getCurrentRow().intValue();
            this.currentColumn = spreadSheetAcessor.getCurrentColumn().intValue();
            
            if(this.rowsThreshold >= 0) {
                this.rowsThreshold += spreadSheetAcessor.getExpandedRows();
            }
            this.expandedRows += spreadSheetAcessor.getExpandedRows();
        }
    }

    public static void write(
        Optional<String> noDataLabel,
        String templateName,
        Object dataSource
    ) {

        try {
            ObjectAccessor objectAccessor = new ObjectAccessor(dataSource);
            XSSFWorkbook myWorkBook = new XSSFWorkbook(SpreadSheetWriter.class.getClassLoader().getResourceAsStream(templateName));
            processSheetNameForVariables(myWorkBook, objectAccessor);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            SpreadSheetAcessor spreadSheetAcessor = processPageInformation(mySheet);
            processSheet(noDataLabel, mySheet, spreadSheetAcessor, objectAccessor);
            spreadSheetAcessor.reset();
            cloneSheet(myWorkBook, spreadSheetAcessor, objectAccessor);

            FileOutputStream os = new FileOutputStream("out.xlsx");
            myWorkBook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static SpreadSheetAcessor processPageInformation(XSSFSheet mySheet) {
        XSSFCell cell = mySheet.getRow(0).getCell(0);
        String comment = cell.getCellComment().getString().getString();
        SpreadSheetAcessor spreadSheetAcessor = new SpreadSheetAcessor(
            Integer.parseInt(comment.split("\n")[1]),
            Integer.parseInt(comment.split("\n")[2])
        );
        cell.removeCellComment();
        mySheet.getRow(0).removeCell(cell);
        return spreadSheetAcessor;
    }

    private static void processSheetNameForVariables(XSSFWorkbook myWorkBook, ObjectAccessor objectAccessor) {
        String sheetName = myWorkBook.getSheetName(0);
        Pattern pattern = Pattern.compile("\\{([a-zA-Z])([a-zA-Z0-9_]*)\\}");
        Matcher matcher = pattern.matcher(sheetName);
        while(matcher.find())
        {
            String match = matcher.group(0);
            sheetName = sheetName.replace(match,
                objectAccessor.access(match.replace("{", "").replace("}", ""))
                    .map(it -> it.getValue())
                    .orElse(""));
        }
        myWorkBook.setSheetName(0,
            WorkbookUtil.createSafeSheetName(sheetName, '_'));
    }

    private static void processSheet(
        Optional<String> noDataLabel,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor,
        ObjectAccessor objectAccessor
    ) {
        while(!spreadSheetAcessor.isDone()) {
            processCell(noDataLabel, mySheet, spreadSheetAcessor, objectAccessor);
            spreadSheetAcessor.next();
        }
    }

    private static void cloneSheet(XSSFWorkbook myWorkBook, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {
        String sheetName = myWorkBook.getSheetName(0);
        myWorkBook.createSheet();
        while(!spreadSheetAcessor.isDone()) {
            copySpreadSheetCell(
                spreadSheetAcessor.getCurrentRow(),
                spreadSheetAcessor.getCurrentColumn(),
                0,
                1,
                myWorkBook
            );
            spreadSheetAcessor.next();
        }
        cloneMergedRegions(myWorkBook);
        myWorkBook.removeSheetAt(0);
        myWorkBook.setSheetName(0, sheetName);
    }

    private static void cloneMergedRegions(XSSFWorkbook myWorkBook) {
        for(CellRangeAddress mergedRegion: myWorkBook.getSheetAt(0).getMergedRegions()) {
            myWorkBook.getSheetAt(1).addMergedRegion(mergedRegion);
        }
    }

    private static void processCell(
        Optional<String> noDataLabel,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor,
        ObjectAccessor objectAccessor
    ) {
        XSSFRow row = mySheet.getRow(spreadSheetAcessor.getCurrentRow());
        if(row != null) {
            XSSFCell cell = row.getCell(spreadSheetAcessor.getCurrentColumn());
            if(cell != null) {
                processCellComment(noDataLabel, cell, mySheet, spreadSheetAcessor, objectAccessor);
                processCellValueForConstants(cell);
                processCellValueForVariables(cell, objectAccessor);
            }
        }
    }

    private static void processCellValueForConstants(XSSFCell cell) {
        if (CellType.STRING.equals(cell.getCellType())) {
            String cellValue = cell.getStringCellValue();
            Pattern pattern = Pattern.compile("\\[([a-zA-Z])([a-zA-Z0-9_]*)\\]");
            Matcher matcher = pattern.matcher(cellValue);
            while (matcher.find()) {
                String match = matcher.group(0);
                switch (match.replace("[", "").replace("]", "")) {
                    case "TODAY":
                        cellValue = cellValue.replace(match,
                            LocalDateTime.now()
                                .format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss")));
                        break;
                }
            }
            cell.setCellValue(cellValue);
        }
    }

    private static void processCellValueForVariables(XSSFCell cell, ObjectAccessor objectAccessor) {
        if (CellType.STRING.equals(cell.getCellType())) {
            String cellValue = cell.getStringCellValue();
            Pattern pattern = Pattern.compile("\\{([a-zA-Z])([a-zA-Z0-9_]*\\?{0,1}.*\\|{0,1}[a-zA-Z0-9_]*)\\}");
            Matcher matcher = pattern.matcher(cellValue);
            while (matcher.find()) {
                String match = matcher.group(0);
                MatchToStringParser parser = new MatchToStringParser(match);
                String value = parser.getValue();
                Optional<String> pipe = parser.getPipe();
                if (pipe.isPresent()) {
                    switch (pipe.get()) {
                        case "DATE":
                            cellValue = cellValue.replace(match,
                                objectAccessor.access(value)
                                    .map(it -> it.getValueFromDate("dd/MMM/yyyy"))
                                    .orElse(parser.getDefaultValue().orElse("")));
                            break;
                        case "INTEGER":
                            cell.setCellValue(objectAccessor.access(value)
                                .map(it -> it.getValueAsInteger())
                                .orElse(0));
                            return;
                        case "DECIMAL":
                            cell.setCellValue(objectAccessor.access(value)
                                .map(it -> it.getValueAsDouble())
                                .orElse(0.0));
                            return;
                    }
                } else {
                    cellValue = cellValue.replace(match,
                        objectAccessor.access(value)
                            .map(it -> it.getValue())
                            .orElse(""));
                }
            }
            cell.setCellValue(cellValue);
        }
    }

    private static void processCellComment(
        Optional<String> noDataLabel,
        XSSFCell cell, XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor,
        ObjectAccessor objectAccessor
    ) {
        XSSFComment comment = cell.getCellComment();
        if(comment != null) {
            String[] lines = comment.getString().getString().split("\n");
            cell.getCellComment().setString("");
            cell.removeCellComment();

            switch(lines[0]) {
                case ":for":
                    processCellCommentFor(noDataLabel,
                        lines[1],
                        Integer.parseInt(lines[2]),
                        Integer.parseInt(lines[3]),
                        mySheet, spreadSheetAcessor, objectAccessor);
                    break;
                case ":mergeBy":
                    processCellCommentMergeBy(lines[1], mySheet, spreadSheetAcessor, objectAccessor);
                    break;
                case ":iterateBy":
                    processCellCommentIterateBy(lines[1], mySheet, spreadSheetAcessor, objectAccessor);
                    break;
            }
        }
    }

    private static void processCellCommentFor(
        Optional<String> noDataLabel,
        String list,
        Integer rows,
        Integer columns,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor,
        ObjectAccessor objectAccessor
    ) {

        objectAccessor.access(list).ifPresent(it -> {
            Integer regionRows = rows;
            Integer regionColumns = columns;

            Integer fromRow = spreadSheetAcessor.getCurrentRow();
            Integer fromColumn = spreadSheetAcessor.getCurrentColumn();
            Integer toRow = spreadSheetAcessor.getCurrentRow();

            if(!noDataLabel.isPresent() || it.listSize() > 0) {
                shiftSpreadSheetRows(fromRow + regionRows,
                    regionRows * (it.listSize() - 1), mySheet, spreadSheetAcessor);

                for(Integer listItem = 0; listItem < it.listSize(); listItem++) {
                    if(listItem > 0) {
                        copySpreadSheetRows(fromRow,
                            fromColumn,
                            toRow,
                            fromColumn,
                            regionRows, regionColumns, mySheet);
                    }
                    toRow += regionRows;
                }
                for(Integer listItem = 0; listItem < it.listSize(); listItem++) {
                    SpreadSheetAcessor subSpreadSheetAcessor = spreadSheetAcessor.SubSpreadSheetAcessor(
                            spreadSheetAcessor.getCurrentRow() + regionRows);

                    processSheet(noDataLabel,
                        mySheet,
                        subSpreadSheetAcessor,
                        it.accessList(listItem).orElseThrow(() ->
                            new NullPointerException()
                        ));

                    spreadSheetAcessor.copyState(subSpreadSheetAcessor);
                }
            } else {
                clearSpreadSheetRow(noDataLabel, fromRow, fromColumn, regionRows, regionColumns, mySheet);
            }
        });
    }

    private static void processCellCommentMergeBy(String list, XSSFSheet mySheet, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {

        objectAccessor.access(list).ifPresent(it -> {
            Integer expandBy = it.listSize() - 1;
            if(expandBy > 0) {
                shiftSpreadSheetColumns(spreadSheetAcessor.getCurrentRow(),
                    spreadSheetAcessor.getCurrentColumn() + 1,
                    expandBy,
                    mySheet,
                    spreadSheetAcessor);

                mySheet.addMergedRegion(
                    new CellRangeAddress(spreadSheetAcessor.getCurrentRow(),
                    spreadSheetAcessor.getCurrentRow(),
                    spreadSheetAcessor.getCurrentColumn(),
                    spreadSheetAcessor.getCurrentColumn() + expandBy)
                );
            }
        });
    }

    private static void processCellCommentIterateBy(String list, XSSFSheet mySheet, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {

        objectAccessor.access(list).ifPresent(it -> {
            Integer iterateBy = it.listSize();
            if(iterateBy > 1) {
                shiftSpreadSheetColumns(
                    spreadSheetAcessor.getCurrentRow(),
                    spreadSheetAcessor.getCurrentColumn() + 1,
                    iterateBy - 1,
                    mySheet,
                    spreadSheetAcessor);
    
                for(int iterate = 1; iterate < iterateBy; iterate++) {
                    copySpreadSheetCell(
                        spreadSheetAcessor.getCurrentRow(),
                        spreadSheetAcessor.getCurrentColumn(),
                        spreadSheetAcessor.getCurrentRow(),
                        spreadSheetAcessor.getCurrentColumn() + iterate,
                        mySheet,
                        false
                    );
                }
            }

            XSSFRow targetRow = mySheet.getRow(spreadSheetAcessor.getCurrentRow());
            for(int iterate = 0; iterate < iterateBy; iterate++) {
                final XSSFCell targetCell = targetRow.getCell(spreadSheetAcessor.getCurrentColumn() + iterate);
                it.accessList(iterate).ifPresent(item -> {
                    processCellValueForVariables(targetCell, item);
                });
            }
        });
    }

    private static void shiftSpreadSheetRows(
        Integer fromRow,
        Integer byRows,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor) {

        if(byRows > 0) {
            for(Integer row = spreadSheetAcessor.getRows() - 1; row >= fromRow; row--) {
                for(Integer column = 0; column < spreadSheetAcessor.getColumns(); column++) {
                    shiftSpreadSheetMergedRegion(
                        row,
                        column,
                        byRows,
                        0,
                        mySheet
                    );
                    copySpreadSheetCell(
                        row,
                        column,
                        row + byRows,
                        column,
                        mySheet,
                        true
                    );
                }
            }
            spreadSheetAcessor.expandRows(byRows);
        }
    }

    private static void shiftSpreadSheetColumns(
        Integer fromRow,
        Integer fromColumn,
        Integer byColumns,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor) {

        for(Integer column = spreadSheetAcessor.getColumns() - 1; column >= fromColumn; column--) {
            shiftSpreadSheetMergedRegion(
                fromRow,
                column,
                0,
                byColumns,
                mySheet
            );
            copySpreadSheetCell(
                fromRow,
                column,
                fromRow,
                column + byColumns,
                mySheet,
                true
            );
        }
        spreadSheetAcessor.expandColumns(byColumns);
    }

    private static void shiftSpreadSheetMergedRegion(
        Integer row,
        Integer column,
        Integer byRows,
        Integer byColumns,
        XSSFSheet mySheet
    ) {
        for(CellRangeAddress mergedRegion: mySheet.getMergedRegions()) {
            if(mergedRegion.getFirstRow() == row &&
                mergedRegion.getFirstColumn() == column) {

                mySheet.removeMergedRegion(mySheet.getMergedRegions().indexOf(mergedRegion));
                mergedRegion.setFirstRow(mergedRegion.getFirstRow() + byRows);
                mergedRegion.setLastRow(mergedRegion.getLastRow() + byRows);
                mergedRegion.setFirstColumn(mergedRegion.getFirstColumn() + byColumns);
                mergedRegion.setLastColumn(mergedRegion.getLastColumn() + byColumns);
                mySheet.addMergedRegion(mergedRegion);
            }
        }
    }

    private static void copySpreadSheetRows(
        Integer fromRow,
        Integer fromColumn,
        Integer toRow,
        Integer toColumn,
        Integer rowsToCopy,
        Integer columnsToCopy,
        XSSFSheet mySheet) {

        for(Integer row = fromRow; row < fromRow + rowsToCopy; row++) {
            for(Integer column = fromColumn; column < fromColumn + columnsToCopy; column++) {
                copySpreadSheetCell(
                    row,
                    column,
                    row + (toRow - fromRow),
                    column + (toColumn - fromColumn),
                    mySheet,
                    false
                );
            }
        }
    }

    private static void copySpreadSheetCell(
        Integer fromRow,
        Integer fromColumn,
        Integer toRow,
        Integer toColumn,
        XSSFSheet mySheet,
        Boolean deleteSource
    ) {

        XSSFRow sourceRow = mySheet.getRow(fromRow);
        if(sourceRow != null) {

            XSSFRow destinationRow = mySheet.getRow(toRow);
            if(destinationRow == null) {
                destinationRow = mySheet.createRow(toRow);
            }

            destinationRow.setHeight(
                sourceRow.getHeight()
            );

            XSSFCell sourceCell = sourceRow.getCell(fromColumn);

            if(sourceCell != null) {

                XSSFCell destinationCell = destinationRow.getCell(toColumn);
                if(destinationCell == null) {
                    destinationCell = destinationRow.createCell(toColumn);
                }

                mySheet.setColumnWidth(toColumn, mySheet.getColumnWidth(fromColumn));

                copySpreadSheetCell(sourceCell, destinationCell, mySheet);

                if(deleteSource) {
                    if(sourceCell.getCellComment() != null) {
                        sourceCell.getCellComment().setString("");
                        sourceCell.removeCellComment();
                    }
                    sourceRow.removeCell(sourceCell);
                }
            }
        }
    }

    private static void copySpreadSheetCell(
        Integer row,
        Integer column,
        Integer fromSheet,
        Integer toSheet,
        XSSFWorkbook myWorkBook
    ) {

        XSSFSheet sourceSheet = myWorkBook.getSheetAt(fromSheet);
        XSSFRow sourceRow = sourceSheet.getRow(row);
        if(sourceRow != null) {

            XSSFSheet destinationSheet = myWorkBook.getSheetAt(toSheet);
            XSSFRow destinationRow = destinationSheet.getRow(row);
            if(destinationRow == null) {
                destinationRow = destinationSheet.createRow(row);
            }

            destinationRow.setHeight(
                sourceRow.getHeight()
            );

            XSSFCell sourceCell = sourceRow.getCell(column);
            if(sourceCell != null) {

                XSSFCell destinationCell = destinationRow.getCell(column);
                if(destinationCell == null) {
                    destinationCell = destinationRow.createCell(column);
                }

                destinationSheet.setColumnWidth(column, sourceSheet.getColumnWidth(column));

                copySpreadSheetCell(sourceCell, destinationCell, destinationSheet);
            }
        }
    }

    private static void copySpreadSheetCell(
        XSSFCell sourceCell,
        XSSFCell destinationCell,
        XSSFSheet destinationSheet
    ) {
        if(CellType.NUMERIC.equals(sourceCell.getCellType())) {
            destinationCell.setCellValue(
                sourceCell.getNumericCellValue()
            );
        } else {
            destinationCell.setCellValue(
                sourceCell.getStringCellValue()
            );            
        }

        destinationCell.setCellStyle(
            sourceCell.getCellStyle()
        );

        if(sourceCell.getCellComment() != null) {
            XSSFComment comment = destinationSheet.createDrawingPatriarch().createCellComment(new XSSFClientAnchor());
            comment.setString(sourceCell.getCellComment().getString().getString());
            destinationCell.setCellComment(comment);
        }
    }

    private static void clearSpreadSheetRow(
        Optional<String> noDataLabel,
        Integer fromRow,
        Integer fromColumn,
        Integer rowsToClear,
        Integer columnsToClear,
        XSSFSheet mySheet
    ) {

        for(Integer row = fromRow; row < fromRow + rowsToClear; row++) {

            XSSFRow cellsRow = mySheet.getRow(row);
            if(row == fromRow && noDataLabel.isPresent()) {
                if(cellsRow == null) {
                    cellsRow = mySheet.createRow(row);
                }
        
                XSSFCell cell = cellsRow.getCell(fromColumn);
                if(cell == null) {
                    cell = cellsRow.createCell(fromColumn);
                }
        
                cell.setCellValue(noDataLabel.get());
            }

            if(cellsRow != null) {
                for(Integer column = fromColumn + ((row == fromRow && noDataLabel.isPresent()) ? 1 : 0); column < fromColumn + columnsToClear; column++) {
                    XSSFCell cell = cellsRow.getCell(column);
                    if(cell != null) {
                        cell.setBlank();
                        if(cell.getCellComment() != null) {
                            cell.getCellComment().setString("");
                            cell.removeCellComment();
                        }
                    }
                }
            }
        }
    }
}
