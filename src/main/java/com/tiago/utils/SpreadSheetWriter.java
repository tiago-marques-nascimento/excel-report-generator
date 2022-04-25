package com.tiago.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadSheetWriter {

    private class SpreadSheetAcessor {
        private Integer rows;
        private Integer columns;
        private Integer currentRow;
        private Integer currentColumn;
        private Integer rowsThreshold;
        private Integer expandedRows;

        public SpreadSheetAcessor(final Integer rows, final Integer columns) {

            this.rows = rows.intValue();
            this.columns = columns.intValue();
            this.currentRow = 0;
            this.currentColumn = 0;
            this.rowsThreshold = -1;
            this.expandedRows = 0;
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

        public Boolean isDone() {
            return this.currentRow >= this.rows ||
                ((this.rowsThreshold >= 0) && this.currentRow >= this.rowsThreshold);
        }

        public void next() {
            if(!this.isDone() && ++this.currentColumn == this.columns) {
                this.currentColumn = 0;
                this.currentRow++;
            }
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

    public void write(String templateName, Object dataSource) {

        try {
            XSSFWorkbook myWorkBook = new XSSFWorkbook(getClass().getClassLoader().getResourceAsStream(templateName));
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            ObjectAccessor objectAccessor = new ObjectAccessor(dataSource);
            SpreadSheetAcessor spreadSheetAcessor = processPageInformation(mySheet);
            processSheet(mySheet, spreadSheetAcessor, objectAccessor);

            FileOutputStream os = new FileOutputStream("out.xlsx");
            myWorkBook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private SpreadSheetAcessor processPageInformation(XSSFSheet mySheet) {
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

    private void processSheet(XSSFSheet mySheet, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {
        while(!spreadSheetAcessor.isDone()) {
            this.processCell(mySheet, spreadSheetAcessor, objectAccessor);
            spreadSheetAcessor.next();
        }
    }

    private void processCell(XSSFSheet mySheet, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {
        XSSFRow row = mySheet.getRow(spreadSheetAcessor.getCurrentRow());
        if(row != null) {
            XSSFCell cell = row.getCell(spreadSheetAcessor.getCurrentColumn());
            if(cell != null) {
                this.processCellComment(cell, mySheet, spreadSheetAcessor, objectAccessor);
                this.processCellValueForConstants(cell, objectAccessor);
                this.processCellValueForVariables(cell, objectAccessor);
            }
        }
    }

    private void processCellValueForConstants(XSSFCell cell, ObjectAccessor objectAccessor) {
        String cellValue = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("\\[([a-zA-Z])([a-zA-Z0-9_]*)\\]");
        Matcher matcher = pattern.matcher(cellValue);
        while(matcher.find())
        {
            String match = matcher.group(0);
            switch(match.replace("[", "").replace("]", "")) {
                case "TODAY":
                    cell.setCellValue(cellValue.replace(match,
                        LocalDateTime.now()
                            .format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"))));
                    break;
            }
        }
    }

    private void processCellValueForVariables(XSSFCell cell, ObjectAccessor objectAccessor) {
        String cellValue = cell.getStringCellValue();
        Pattern pattern = Pattern.compile("\\{([a-zA-Z])([a-zA-Z0-9_]*)\\}");
        Matcher matcher = pattern.matcher(cellValue);
        while(matcher.find())
        {
            String match = matcher.group(0);
            objectAccessor.access(
                match.replace("{", "").replace("}", ""))
                    .ifPresent(it -> {
                        cell.setCellValue(cellValue.replace(match, it.getValue()));
                    });
        }
    }

    private void processCellComment(XSSFCell cell, XSSFSheet mySheet, SpreadSheetAcessor spreadSheetAcessor, ObjectAccessor objectAccessor) {
        XSSFComment comment = cell.getCellComment();
        if(comment != null) {
            String[] lines = comment.getString().getString().split("\n");
            cell.getCellComment().setString("");
            cell.removeCellComment();

            switch(lines[0]) {
                case ":for":
                    objectAccessor.access(lines[1]).ifPresent(it -> {
                        Integer regionRows = Integer.parseInt(lines[2]);
                        Integer regionColumns = Integer.parseInt(lines[3]);

                        this.shiftSpreadSheetRows(spreadSheetAcessor.getCurrentRow() + regionRows,
                            regionRows * (it.listSize() - 1), mySheet, spreadSheetAcessor);

                        Integer fromRow = spreadSheetAcessor.getCurrentRow();
                        Integer fromColumn = spreadSheetAcessor.getCurrentColumn();
                        Integer toRow = spreadSheetAcessor.getCurrentRow();

                        for(Integer listItem = 0; listItem < it.listSize(); listItem++) {
                            if(listItem > 0) {
                                this.copySpreadSheetRows(fromRow,
                                    fromColumn,
                                    toRow,
                                    fromColumn,
                                    regionRows, regionColumns, mySheet, spreadSheetAcessor);
                            }
                            toRow += regionRows;
                        }

                        for(Integer listItem = 0; listItem < it.listSize(); listItem++) {
                            SpreadSheetAcessor subSpreadSheetAcessor = spreadSheetAcessor.SubSpreadSheetAcessor(
                                    spreadSheetAcessor.getCurrentRow() + regionRows);

                            processSheet(mySheet,
                                subSpreadSheetAcessor,
                                it.accessList(listItem).orElseThrow(() ->
                                    new NullPointerException()
                                ));

                            spreadSheetAcessor.copyState(subSpreadSheetAcessor);
                        }
                    });
                    break;
            }
        }
    }

    private void shiftSpreadSheetRows(
        Integer fromRow,
        Integer rows,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor) {

        for(Integer row = spreadSheetAcessor.getRows() - 1; row >= fromRow; row--) {
            for(Integer column = 0; column < spreadSheetAcessor.getColumns(); column++) {
                this.copySpreadSheetCell(
                    row,
                    column,
                    row + rows,
                    column,
                    mySheet,
                    true
                );
            }
        }
        spreadSheetAcessor.expandRows(rows);
    }

    private void copySpreadSheetRows(
        Integer fromRow,
        Integer fromColumn,
        Integer toRow,
        Integer toColumn,
        Integer rowsToCopy,
        Integer columnsToCopy,
        XSSFSheet mySheet,
        SpreadSheetAcessor spreadSheetAcessor) {

        for(Integer row = fromRow; row < fromRow + rowsToCopy; row++) {
            for(Integer column = fromColumn; column < fromColumn + columnsToCopy; column++) {
                this.copySpreadSheetCell(
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

    private void copySpreadSheetCell(
        Integer fromRow,
        Integer fromColumn,
        Integer toRow,
        Integer toColumn,
        XSSFSheet mySheet,
        Boolean deleteSource
    ) {

        XSSFRow sourceRow = mySheet.getRow(fromRow);
        if(sourceRow != null) {
            XSSFCell sourceCell = sourceRow.getCell(fromColumn);
            if(sourceCell != null) {

                XSSFRow destinationRow = mySheet.getRow(toRow);
                if(destinationRow == null) {
                    destinationRow = mySheet.createRow(toRow);
                }

                XSSFCell destinationCell = destinationRow.getCell(toColumn);
                if(destinationCell == null) {
                    destinationCell = destinationRow.createCell(toColumn);
                }

                destinationCell.setCellValue(
                    sourceCell.getStringCellValue()
                );

                destinationCell.setCellStyle(
                    sourceCell.getCellStyle()
                );

                if(sourceCell.getCellComment() != null) {
                    XSSFComment comment = mySheet.createDrawingPatriarch().createCellComment(new XSSFClientAnchor());
                    comment.setString(sourceCell.getCellComment().getString().getString());
                    destinationCell.setCellComment(comment);
                }

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
}
