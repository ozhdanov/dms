package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.model.RenaissanceModel;
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
public class Renaissance extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${renaissance.storage}")
    private String storageFileUrl;
    @Value("${renaissance.sender}")
    private String senderEmailTemplate;
    @Value("${renaissance.liststemplate}")
    private String listsTemplate;
    @Value("${renaissance.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${renaissance.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${renaissance.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${renaissance.deattachfiletemplate}")
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

    public List<RenaissanceModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<RenaissanceModel> customers = new ArrayList<>();
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
                        if (cell.toString().contains("на срок:")) {
                            cell = cellIterator.next();
                            validity = cell.toString();
                            break;
                        } else if (cell.toString().contains("сотрудников:")) {
                            cell = cellIterator.next();
                            placeOfWork = cell.toString();
                            break;
                        } else if (cell.toString().contains("№ п/п")) {
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
                            log.debug("Прикрепление пациента");
                            RenaissanceModel renaissanceModel = new RenaissanceModel();

                            renaissanceModel.setFio(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setDateOfBirth(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setPassport(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setAdress(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setPhoneNumber(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setPolicyNumber(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            renaissanceModel.setValidity(validity);
                            renaissanceModel.setPlaceOfWork(placeOfWork);

                            log.debug(renaissanceModel.toString());

                            customers.add(renaissanceModel);
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

    public List<RenaissanceModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<RenaissanceModel> customers = new ArrayList<>();
        try {
            // we create an XSSF Workbook object for our XLSX Excel File
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
                            log.debug("Открепление пациента");

                            RenaissanceModel renaissanceModel = new RenaissanceModel();

                            renaissanceModel.setPolicyNumber(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());


                            log.debug(renaissanceModel.toString());

                            customers.add(renaissanceModel);
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

    public void addCustomersToFile(List<RenaissanceModel> customers) {

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
                    Cell validityCell = row.getCell(6);
                    validityCell.setCellType(CellType.STRING);
                    String validity = validityCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (RenaissanceModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()) && validity.equals(customer.getValidity()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (RenaissanceModel customer : customers) {
                if (customer.isNew()) {

                    log.info("Прикрепление пациента {}", customer.getFio());

                    int rows = sheet.getPhysicalNumberOfRows() - sheet.getFirstRowNum();
                    sheet.shiftRows(1,rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getFio());
                    row.createCell(2).setCellValue(customer.getPassport());
                    row.createCell(3).setCellValue(customer.getDateOfBirth());
                    row.createCell(4).setCellValue(customer.getAdress());
                    row.createCell(5).setCellValue(customer.getPhoneNumber());
                    row.createCell(6).setCellValue(customer.getValidity());
                    row.createCell(7).setCellValue(customer.getPlaceOfWork());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Ренессанс", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

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


    public void removeCustomersFromFile(List<RenaissanceModel> customers) {
        for (RenaissanceModel customer : customers) {
            removeCustomerFromFile("Ренессанс", storageFileUrl, customer.getPolicyNumber(), 0);
        }
    }

}
