package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.model.SoglasieModel;
import oz.med.DMSParser.services.EmailService;

import java.awt.*;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
@Slf4j
public class Soglasie extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${soglasie.storage}")
    private String storageFileUrl;
    @Value("${soglasie.sender}")
    private String senderEmailTemplate;
    @Value("${soglasie.liststemplate}")
    private String listsTemplate;
    @Value("${soglasie.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${soglasie.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${soglasie.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${soglasie.deattachfiletemplate}")
    private String deattachFileTemplate;

    public boolean isListsMail(String from, String subject) {
        return this.isListsMail(from, subject, senderEmailTemplate, listsTemplate);
    }
    public boolean isAttachListMail(String from, String subject) {
        return this.isAttachListMail(from, subject, senderEmailTemplate, attachListTemplate);
    }
    public boolean isDeattachListMail(String from, String subject) {
        return this.isDeattachListsMail(from, subject, senderEmailTemplate, deattachListTemplate);
    }
    public boolean isAttachFile(String fileName) {
        return this.isAttachFile(fileName, attachFileTemplate);
    }
    public boolean isDeattachFile(String fileName) {
        return this.isDeattachFile(fileName, deattachFileTemplate);
    }

    public List<SoglasieModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<SoglasieModel> customers = new ArrayList<>();
        try {
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new HSSFWorkbook(is);
            // we get first sheet
            HSSFSheet sheet = workbook.getSheetAt(0);

            // we iterate on rows
            Iterator<Row> rowIt = sheet.iterator();

            boolean startOfDataFlag = false;
            int prewRowIndex = 0;
            String validity = null;
            String placeOfWork = null;

            while (rowIt.hasNext()) {
                Row row = rowIt.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Cell cell;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        //Вытаскиваем сроки страхования
                        if (cell.toString().contains("Прикрепление")) {
                            cell = cellIterator.next();
                            validity = cell.toString();
                            cellIterator.next();
                            cellIterator.next();
                            cell = cellIterator.next();
                            validity += " - " + cell.toString();
                            break;
                        } else if (cell.toString().contains("Организация")) {
                            cell = cellIterator.next();
                            placeOfWork = cell.toString();
                            break;
                        } else if (cell.toString().contains("№ полиса ДМС")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cellIterator.next();
                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        continue;
                    }

                    if(cellIterator.hasNext() && !cell.toString().isEmpty()) {
                        try {
                            log.info("Прикрепление пациента");
                            SoglasieModel soglasieModel = new SoglasieModel();

                            soglasieModel.setSurname(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setName(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setPatronymic(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setDateOfBirth(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setAdress(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setPhoneNumber(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setPolicyNumber(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            soglasieModel.setValidity(validity);
                            soglasieModel.setPlaceOfWork(placeOfWork);

                            log.info(soglasieModel.toString());

                            customers.add(soglasieModel);
                        } catch (Exception e) {
                            log.error("Ошибка парсинга строки", e);
                        }
                    } else{
                        startOfDataFlag = false;
                    }

                }

                if(startOfDataFlag)
                    prewRowIndex = row.getRowNum();

            }

            workbook.close();
            is.close();
        } catch (IOException e) {
            log.error("Не удалось распарсить документ", e);
        } finally {
            try {
                workbook.close();
                is.close();
            } catch (Exception e) {
            }
        }

        return customers;

    }

    public List<SoglasieModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<SoglasieModel> customers = new ArrayList<>();
        try {
            // we create an HSSF Workbook object for our XLSX Excel File
            workbook = new HSSFWorkbook(is);
            // we get first sheet
            HSSFSheet sheet = workbook.getSheetAt(0);

            // we iterate on rows
            Iterator<Row> rowIt = sheet.iterator();

            boolean startOfDataFlag = false;

            int prewRowIndex = 0;

            while (rowIt.hasNext()) {
                Row row = rowIt.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Cell cell = null;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        if (cell.toString().equals("Нормер полиса")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        continue;
                    }

                    if(cellIterator.hasNext() && !cell.toString().isEmpty()) {
                        try {
                            log.info("Открепление пациента");

                            SoglasieModel soglasieModel = new SoglasieModel();

                            soglasieModel.setPolicyNumber(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());


                            log.info(soglasieModel.toString());

                            customers.add(soglasieModel);
                        } catch (Exception e) {
                            log.error("Ошибка парсинга строки", e);
                        }
                    }

                }
                prewRowIndex = row.getRowNum();
            }

            workbook.close();
            is.close();
        } catch (IOException e) {
            log.error("Не удалось распарсить документ", e);
        } finally {
            try {
                workbook.close();
                is.close();
            } catch (Exception e) {
            }
        }

        return customers;

    }

    public void addCustomersToFile(List<SoglasieModel> customers) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(this.listsUrl + storageFileUrl));
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(inputStream);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() > 0 && !isRowEmpty(row)) {
                    Cell policyNumberCell = row.getCell(0);
                    String policyNumber = policyNumberCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (SoglasieModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (SoglasieModel customer : customers) {
                if (customer.isNew()) {
                    int rows = sheet.getLastRowNum();
                    sheet.shiftRows(1,rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getSurname());
                    row.createCell(2).setCellValue(customer.getName());
                    row.createCell(3).setCellValue(customer.getPatronymic());
                    row.createCell(4).setCellValue(customer.getDateOfBirth());
                    row.createCell(5).setCellValue(customer.getAdress());
                    row.createCell(6).setCellValue(customer.getPhoneNumber());
                    row.createCell(7).setCellValue(customer.getValidity());
                    row.createCell(8).setCellValue(customer.getPlaceOfWork());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Согласие", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

            FileOutputStream outputStream = new FileOutputStream(this.listsUrl + storageFileUrl);
            workbook.write(outputStream);
            outputStream.close();

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


    public void removeCustomersFromFile(List<SoglasieModel> customers) {
        for (SoglasieModel customer : customers) {
            removeCustomerFromFile("Согласие", storageFileUrl, customer.getPolicyNumber(), 0);
        }
    }

}
