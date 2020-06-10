package oz.med.DMSParser.companies;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

@Service
public class Company {

    public boolean isListsMail(String from, String subject, String senderEmailTemplate, String listsTemplate) {
        if (from.toUpperCase().contains(senderEmailTemplate.toUpperCase()) && subject.toUpperCase().contains(listsTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isAttachListMail(String from, String subject, String senderEmailTemplate, String attachListsTemplate) {
        if (from.toUpperCase().contains(senderEmailTemplate.toUpperCase()) && subject.toUpperCase().contains(attachListsTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isDeattachListsMail(String from, String subject, String senderEmailTemplate, String deattachListsTemplate) {
        if (from.toUpperCase().contains(senderEmailTemplate.toUpperCase()) && subject.toUpperCase().contains(deattachListsTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isAttachFile(String fileName, String attacFileTemplate) {
        if (fileName.toUpperCase().contains(attacFileTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isDeattachFile(String fileName, String deattachFileTemplate) {
        if (fileName.toUpperCase().contains(deattachFileTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public void removeExcelRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

//    public boolean isRowEmpty(Row row) {
//        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
//            Cell cell = row.getCell(c);
//            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
//                return false;
//        }
//        return true;
//    }

    public boolean isRowEmpty(Row row) {
        boolean isEmpty = true;
        DataFormatter dataFormatter = new DataFormatter();

        if (row != null) {
            for (Cell cell : row) {
                if (dataFormatter.formatCellValue(cell).trim().length() > 0) {
                    isEmpty = false;
                    break;
                }
            }
        }

        return isEmpty;
    }

}
