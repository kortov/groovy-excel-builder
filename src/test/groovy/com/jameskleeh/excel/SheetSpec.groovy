package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import spock.lang.Specification

class SheetSpec extends Specification {

    void "test skipRows"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row()
                skipRows(2)
                row()
            }
        }

        when:
        Iterator<Row> rows = workbook.getSheetAt(0).rowIterator()

        then:
        rows.next().rowNum == 0
        rows.next().rowNum == 3
        !rows.hasNext()
    }

    void "test row(Object...)"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row(1, 2, 3)
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.physicalNumberOfCells == 3
        row.getCell(0).numericCellValue == 1
        row.getCell(1).numericCellValue == 2
        row.getCell(2).numericCellValue == 3
    }

    void "test row(Map options)"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet {
                row([height: 12F]) {

                }
            }
        }

        when:
        Row row = workbook.getSheetAt(0).getRow(0)

        then:
        row.heightInPoints == 12F
    }

    void "test sheet() and autosize column without tracking it"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet() {
                row(1)
            }
        }

        when:
        workbook.getSheetAt(0).autoSizeColumn(0)

        then:
        final IllegalStateException exception = thrown()
        exception.message == 'Could not auto-size column. Make sure the column was tracked prior to auto-sizing the column.'
    }

    void "test sheet([trackColumns: int]) and autosize column"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: 0]) {
                row(1)
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.trackedColumnsForAutoSizing == ([0] as Set)
        sheet.autoSizeColumn(0)
    }

    void "test sheet([trackColumns: [int, int]])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: [1, 3]]) {
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.trackedColumnsForAutoSizing == ([1, 3] as Set)
    }

    void "test sheet([trackColumns: TRACK_ALL])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: AutoSizeColumnsTracking.TRACK_ALL]) {
                row(1, 2)
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
//         can't use getTrackedColumnsForAutoSizing(). If all columns are tracked, that
//         will only return the columns that have been explicitly or implicitly tracked

//         this method is buggy (fixed in https://github.com/apache/poi/pull/175)
//         but for the test it's not a big problem
        sheet.isColumnTrackedForAutoSizing(1)
    }

    void "test sheet([trackColumns: [0], untrackColumns: 0])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: [0],
                   untrackColumns: 0]) {
                row(1, 2)
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        !sheet.isColumnTrackedForAutoSizing(0)
    }

    void "test reversed order sheet([untrackColumns: [0], trackColumns: 0])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([untrackColumns: [0],
                   trackColumns: 0]) {
                row(1, 2)
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        !sheet.isColumnTrackedForAutoSizing(0)
    }

    void "test sheet([trackColumns: [int, int], untrackColumns: [int, int]])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: [0, 1, 2],
                   untrackColumns: [1, 2]]) {
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        sheet.trackedColumnsForAutoSizing == ([0] as Set)
    }

    void "test sheet([trackColumns: TRACK_ALL, untrackColumns: TRACK_ALL])"() {
        SXSSFWorkbook workbook = ExcelBuilder.build {
            sheet([trackColumns: AutoSizeColumnsTracking.TRACK_ALL,
                   untrackColumns: AutoSizeColumnsTracking.UNTRACK_ALL]) {
                row(1)
            }
        }

        when:
        SXSSFSheet sheet = workbook.getSheetAt(0)

        then:
        // can't use getTrackedColumnsForAutoSizing(). If all columns are tracked, that
        // will only return the columns that have been explicitly or implicitly tracked

        // this method is buggy (fixed in https://github.com/apache/poi/pull/175)
        // but for the test it's not a big problem
        !sheet.isColumnTrackedForAutoSizing(0)
    }

    void "test sheet([untrackColumns: float])"() {
        when:
        ExcelBuilder.build {
            sheet([untrackColumns: 12F]) {
            }
        }

        then:
        final IllegalArgumentException exception = thrown()
        exception.message == 'Sheet untracked columns must be integers'
    }
}
