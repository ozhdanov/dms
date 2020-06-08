package oz.med.DMSParser.companies;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

@Service
public class Company {

    public boolean isListsMail(String from, String subject, String senderEmailTemplate, String listsTemplate) {
        if (from.contains(senderEmailTemplate) && subject.toUpperCase().contains(listsTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isAttachListMail(String from, String subject, String senderEmailTemplate, String attachListsTemplate) {
        if (from.contains(senderEmailTemplate) && subject.toUpperCase().contains(attachListsTemplate.toUpperCase()))
            return true;
        else
            return false;
    }

    public boolean isDeattachListsMail(String from, String subject, String senderEmailTemplate, String deattachListsTemplate) {
        if (from.contains(senderEmailTemplate) && subject.toUpperCase().contains(deattachListsTemplate.toUpperCase()))
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

}
