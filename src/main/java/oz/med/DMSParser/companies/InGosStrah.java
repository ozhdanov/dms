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
import oz.med.DMSParser.model.InGosStrahModel;
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
public class InGosStrah extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${ingosstrah.storage}")
    private String storageFileUrl;
    @Value("${ingosstrah.sender}")
    private String senderEmailTemplate;
    @Value("${ingosstrah.liststemplate}")
    private String listsTemplate;
    @Value("${ingosstrah.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${ingosstrah.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${ingosstrah.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${ingosstrah.deattachfiletemplate}")
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

    public List<InGosStrahModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<InGosStrahModel> customers = new ArrayList<>();
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
                        if (cell.toString().equals("п/п")) {
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
                            log.info("Прикрепление пациента");
                            InGosStrahModel inGosStrahModel = new InGosStrahModel();

                            inGosStrahModel.setPolicyNumber(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setSurname(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setName(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPatronymic(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(format.parse(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setSex(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setAdressAndPhoneNumber(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setNote(row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPlan(row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceProgram(row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyStartDate(format.parse(row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setPolicyEndDate(format.parse(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setInsuranceNote(row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceExtension(row.getCell(14, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setLimitations(row.getCell(15, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setCharacteristic(row.getCell(16, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                            log.info(inGosStrahModel.toString());

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

    public List<InGosStrahModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<InGosStrahModel> customers = new ArrayList<>();
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
                        if (cell.toString().equals("п/п")) {
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
                            log.info("Открепление пациента");

                            InGosStrahModel inGosStrahModel = new InGosStrahModel();

                            inGosStrahModel.setPolicyNumber(row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setSurname(row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setName(row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPatronymic(row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setDateOfBirth(format.parse(row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setSex(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setAdressAndPhoneNumber(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setNote(row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPlan(row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceProgram(row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setPolicyStartDate(format.parse(row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setPolicyEndDate(format.parse(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString()));
                            inGosStrahModel.setInsuranceNote(row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setInsuranceExtension(row.getCell(14, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setLimitations(row.getCell(15, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());
                            inGosStrahModel.setCharacteristic(row.getCell(16, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString());

                            log.info(inGosStrahModel.toString());

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

    public void addCustomersToFile(List<InGosStrahModel> customers) {

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
                        for (InGosStrahModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (InGosStrahModel customer : customers) {
                if (customer.isNew()) {
                    int rows = sheet.getLastRowNum();
                    sheet.shiftRows(1,rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getSurname());
                    row.createCell(2).setCellValue(customer.getName());
                    row.createCell(3).setCellValue(customer.getPatronymic());
                    row.createCell(4).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(5).setCellValue(customer.getSex());
                    row.createCell(6).setCellValue(customer.getAdressAndPhoneNumber());
                    row.createCell(7).setCellValue(customer.getNote());
                    row.createCell(8).setCellValue(customer.getPlan());
                    row.createCell(9).setCellValue(customer.getInsuranceProgram());
                    row.createCell(10).setCellValue(format.format(customer.getPolicyStartDate()));
                    row.createCell(11).setCellValue(format.format(customer.getPolicyEndDate()));
                    row.createCell(12).setCellValue(customer.getInsuranceNote());
                    row.createCell(13).setCellValue(customer.getInsuranceExtension());
                    row.createCell(14).setCellValue(customer.getLimitations());
                    row.createCell(15).setCellValue(customer.getCharacteristic());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Ингосстрах", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

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

    public void removeCustomersFromFile(List<InGosStrahModel> customers) {
        for (InGosStrahModel customer : customers) {
            removeCustomerFromFile("Ингосстрах", storageFileUrl, customer.getPolicyNumber(), 0);
        }
    }

}
