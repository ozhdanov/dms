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
import oz.med.DMSParser.model.AbsolutModel;
import oz.med.DMSParser.model.AlfaStrahModel;
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
public class Absolut extends Company {

    private DateFormat format = new SimpleDateFormat("MM/dd/yyyy");

    @Value("${absolut.storage}")
    private String storageFileUrl;
    @Value("${absolut.sender}")
    private String senderEmailTemplate;
    @Value("${absolut.liststemplate}")
    private String listsTemplate;
    @Value("${absolut.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${absolut.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${absolut.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${absolut.deattachfiletemplate}")
    private String deattachFileTemplate;
    @Value("${absolut.attachlisttemplate2}")
    private String attachListTemplate2;
    @Value("${absolut.deattachlisttemplate2}")
    private String deattachListTemplate2;
    @Value("${absolut.attachfiletemplate2}")
    private String attachFileTemplate2;
    @Value("${absolut.deattachfiletemplate2}")
    private String deattachFileTemplate2;

    public boolean isListsMail(String from, String subject) {
        return this.isListsMail(from, subject, senderEmailTemplate, listsTemplate);
    }
    public boolean isAttachListMail(String from, String subject) {
        return (this.isAttachListMail(from, subject, senderEmailTemplate, attachListTemplate) || this.isAttachListMail(from, subject, senderEmailTemplate, attachListTemplate2));
    }
    public boolean isDeattachListMail(String from, String subject) {
        return (this.isDeattachListsMail(from, subject, senderEmailTemplate, deattachListTemplate) || this.isDeattachListsMail(from, subject, senderEmailTemplate, deattachListTemplate2));
    }
    public boolean isAttachFile(String fileName) {
        return (this.isAttachFile(fileName, attachFileTemplate) || this.isAttachFile(fileName, attachFileTemplate2));
    }
    public boolean isDeattachFile(String fileName) {
        return (this.isDeattachFile(fileName, deattachFileTemplate) || this.isDeattachFile(fileName, deattachFileTemplate2));
    }

    public List<AbsolutModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<AbsolutModel> customers = new ArrayList<>();
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
                            AbsolutModel inGosStrahModel = new AbsolutModel();

                            inGosStrahModel.setFio(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setAdress(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyNumber(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setValidity(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceProgram(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsurant(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

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

    public List<AbsolutModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<AbsolutModel> customers = new ArrayList<>();
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

                            AbsolutModel inGosStrahModel = new AbsolutModel();

                            inGosStrahModel.setFio(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setAdress(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyNumber(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setValidity(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceProgram(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsurant(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

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

    public void addCustomersToFile(List<AbsolutModel> customers) {

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
                    Cell policyNumberCell = row.getCell(3);
                    String policyNumber = policyNumberCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (AbsolutModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (AbsolutModel customer : customers) {
                if (customer.isNew()) {

                    log.info("Прикрепление пациента {}", customer.getFio());

                    int rows = sheet.getPhysicalNumberOfRows() - sheet.getFirstRowNum();
                    sheet.shiftRows(1, rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getFio());
                    row.createCell(1).setCellValue(customer.getDateOfBirth());
                    row.createCell(2).setCellValue(customer.getAdress());
                    row.createCell(3).setCellValue(customer.getPolicyNumber());
                    row.createCell(4).setCellValue(customer.getValidity());
                    row.createCell(5).setCellValue(customer.getInsuranceProgram());
                    row.createCell(6).setCellValue(customer.getInsurant());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Абсолют", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

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

    public void removeCustomersFromFile(List<AbsolutModel> customers) {
        for (AbsolutModel customer : customers) {
            removeCustomerFromFile("Абсолют", storageFileUrl, customer.getPolicyNumber(), 3);
        }
    }

}
