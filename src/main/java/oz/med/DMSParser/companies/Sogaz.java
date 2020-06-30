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
import oz.med.DMSParser.model.SogazModel;
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
public class Sogaz extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${sogaz.storage}")
    private String storageFileUrl;
    @Value("${sogaz.sender}")
    private String senderEmailTemplate;
    @Value("${sogaz.liststemplate}")
    private String listsTemplate;
    @Value("${sogaz.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${sogaz.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${sogaz.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${sogaz.deattachfiletemplate}")
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

    public List<SogazModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<SogazModel> customers = new ArrayList<>();
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

                Cell cell;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        if (cell.toString().equals("№ п/п")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cell = cellIterator.next();

                    //Ожидаем порядковый номер, а встречаем что-то длиньше
                    if (cell.toString().length() > 3) {
                        startOfDataFlag = false;
                        continue;
                    }

                    if(cellIterator.hasNext() && !cell.toString().isEmpty()) {
                        try {
                            log.debug("Прикрепление пациента");
                            SogazModel inGosStrahModel = new SogazModel();

                            inGosStrahModel.setSurname(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setName(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPatronymic(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(format.parse(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setSex(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setAdress(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setHomePhoneNumber(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setWorkPhoneNumber(row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setMobilPhoneNumber(row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyNumber(row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyStartDate(format.parse(row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setPolicyEndDate(format.parse(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setInsuranceProgram(row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPlaceOfWork(row.getCell(14, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPosition(row.getCell(15, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                            log.debug(inGosStrahModel.toString());

                            customers.add(inGosStrahModel);
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

    public List<SogazModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<SogazModel> customers = new ArrayList<>();
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
                        if (cell.toString().equals("№ п/п")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cell = cellIterator.next();

                    //Ожидаем порядковый номер, а встречаем что-то длиньше
                    if (cell.toString().length() > 3) {
                        startOfDataFlag = false;
                        continue;
                    }

                    if(cellIterator.hasNext() && !cell.toString().isEmpty()) {
                        try {
                            log.debug("Открепление пациента");

                            SogazModel inGosStrahModel = new SogazModel();

                            inGosStrahModel.setSurname(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setName(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPatronymic(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(format.parse(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setPolicyNumber(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyEndDate(format.parse(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setInsuranceProgram(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPlaceOfWork(row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                            log.debug(inGosStrahModel.toString());

                            customers.add(inGosStrahModel);
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

    public void addCustomersToFile(List<SogazModel> customers) {

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
                    Cell policyNumberCell = row.getCell(9);
                    String policyNumber = policyNumberCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (SogazModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (SogazModel customer : customers) {
                if (customer.isNew()) {

                    log.info("Прикрепление пациента {} {} {}", customer.getSurname(), customer.getName(), customer.getPatronymic());

                    int rows = sheet.getPhysicalNumberOfRows() - sheet.getFirstRowNum();
                    sheet.shiftRows(1,rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getSurname());
                    row.createCell(1).setCellValue(customer.getName());
                    row.createCell(2).setCellValue(customer.getPatronymic());
                    row.createCell(3).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(4).setCellValue(customer.getSex());
                    row.createCell(5).setCellValue(customer.getAdress());
                    row.createCell(6).setCellValue(customer.getHomePhoneNumber());
                    row.createCell(7).setCellValue(customer.getWorkPhoneNumber());
                    row.createCell(8).setCellValue(customer.getMobilPhoneNumber());
                    row.createCell(9).setCellValue(customer.getPolicyNumber());
                    row.createCell(10).setCellValue(format.format(customer.getPolicyStartDate()));
                    row.createCell(11).setCellValue(format.format(customer.getPolicyEndDate()));
                    row.createCell(12).setCellValue(customer.getInsuranceProgram());
                    row.createCell(13).setCellValue(customer.getPlaceOfWork());
                    row.createCell(14).setCellValue(customer.getPosition());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Согаз", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

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


    public void removeCustomersFromFile(List<SogazModel> customers) {
        for (SogazModel customer : customers) {
            removeCustomerFromFile("Согаз", storageFileUrl, customer.getPolicyNumber(), 9);
        }
    }

}
