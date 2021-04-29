package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.MyTrayIcon;
import oz.med.DMSParser.services.EmailService;

import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
@Slf4j
public class Company {

    @Autowired
    MyTrayIcon myTrayIcon;

    @Value("${listsUrl}")
    public String listsUrl;

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

    public boolean isAttachFile(String fileName, String attachFileTemplate) {
        if (fileName.toUpperCase().contains(attachFileTemplate.toUpperCase()))
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

    public void removeCustomerFromFile(String companyName, String storageFileUrl, String policyNumber, int cellNumber) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileInputStream inputStream = null;
        List<Row> listOfRowsToRemove = new ArrayList<>();
        try {
            inputStream = new FileInputStream(new File(this.listsUrl + storageFileUrl));
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(inputStream);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getCell(cellNumber) != null) row.getCell(cellNumber).setCellType(CellType.STRING);
                if (row.getRowNum() > 0
                        && !isRowEmpty(row)
                        && !row.getCell(cellNumber).getStringCellValue().isEmpty()
                        && row.getCell(cellNumber).getStringCellValue().equals(policyNumber)) {
                    listOfRowsToRemove.add(row);
                    break;
                }
            }

            int currentDeattachCount = 0;
            for(Row row: listOfRowsToRemove){

                log.info("Открепление пациента  {}", policyNumber);

                removeExcelRow(sheet, row.getRowNum());
                EmailService.deattachCount++;
                currentDeattachCount++;
            }
            if (currentDeattachCount > 0)
                myTrayIcon.displayMessage(companyName, "Откреплено " + currentDeattachCount + " пациентов", TrayIcon.MessageType.INFO);

            FileOutputStream outputStream = new FileOutputStream(this.listsUrl + storageFileUrl);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            workbook.close();
            inputStream.close();
        } catch (FileNotFoundException e) {
            log.error("Процесс не может получить доступ к файлу", e);
            myTrayIcon.displayMessage("Ошибка", e.getLocalizedMessage(), TrayIcon.MessageType.ERROR);
        } catch (IOException e) {
            log.error("Не удалось распарсить документ", e);
        } finally {
            try {
                workbook.close();
                inputStream.close();
            } catch (Exception e) {
            }
        }
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
