package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRow;
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
import oz.med.DMSParser.model.InGosStrahModel;
import oz.med.DMSParser.model.RosGosStrahModel;
import oz.med.DMSParser.services.EmailService;

import java.awt.*;
import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@Service
@Slf4j
public class RosGosStrah extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${rosgosstrah.storage}")
    private String storageFileUrl;
    @Value("${rosgosstrah.sender}")
    private String senderEmailTemplate;
    @Value("${rosgosstrah.liststemplate}")
    private String listsTemplate;
    @Value("${rosgosstrah.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${rosgosstrah.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${rosgosstrah.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${rosgosstrah.deattachfiletemplate}")
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

    public List<RosGosStrahModel> parseAttachListExcel(InputStream is) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<RosGosStrahModel> customers = new ArrayList<>();
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

            while (rowIt.hasNext()) {
                Row row = rowIt.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Cell cell;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        //Вытаскиваем сроки страхования
                        if (cell.toString().contains("Срок страхования")) {
                            validity = cell.toString().substring(cell.toString().indexOf("\n") + 1, cell.toString().indexOf("Страхователь") - 1);
                            break;
                        } else if (cell.toString().equals("№ п/п")) {
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
//                            log.debug("Прикрепление пациента");
                            RosGosStrahModel rosGosStrahModel = new RosGosStrahModel();

                            rosGosStrahModel.setFio(cellIterator.next().toString());
                            rosGosStrahModel.setSex(cellIterator.next().toString());
                            rosGosStrahModel.setDateOfBirth(cellIterator.next().getDateCellValue());
                            rosGosStrahModel.setAdress(cellIterator.next().toString());
                            rosGosStrahModel.setPhoneNumber(cellIterator.next().toString());
                            rosGosStrahModel.setValidity(validity);

                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(CellType.STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();
                            rosGosStrahModel.setPolicyNumber(policyNumber);

//                            log.debug(rosGosStrahModel.toString());

                            customers.add(rosGosStrahModel);
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

    public List<RosGosStrahModel> parseDeattachListExcel(InputStream is) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        List<RosGosStrahModel> customers = new ArrayList<>();
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
//                            log.debug("Открепление пациента");

                            RosGosStrahModel rosGosStrahModel = new RosGosStrahModel();

                            rosGosStrahModel.setFio(cellIterator.next().toString());
                            rosGosStrahModel.setSex(cellIterator.next().toString());
                            rosGosStrahModel.setDateOfBirth(cellIterator.next().getDateCellValue());
                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(CellType.STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();
                            rosGosStrahModel.setPolicyNumber(policyNumber);

//                            log.debug(rosGosStrahModel.toString());

                            customers.add(rosGosStrahModel);
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

    public void addCustomersToFile(List<RosGosStrahModel> customers) {

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
                    Cell policyNumberCell = row.getCell(5);
                    String policyNumber = policyNumberCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (RosGosStrahModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            int currentAttachCount = 0;
            for (RosGosStrahModel customer : customers) {
                if (customer.isNew()) {

                    log.info("Прикрепление пациента {}", customer.getFio());

                    int rows = sheet.getPhysicalNumberOfRows() - sheet.getFirstRowNum();
                    sheet.shiftRows(1,rows,1);
                    XSSFRow row = sheet.createRow(1);
                    row.createCell(0).setCellValue(customer.getFio());
                    row.createCell(1).setCellValue(customer.getSex());
                    if(customer.getDateOfBirth() != null) row.createCell(2).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(3).setCellValue(customer.getAdress());
                    row.createCell(4).setCellValue(customer.getPhoneNumber());
                    row.createCell(5).setCellValue(customer.getPolicyNumber());
                    row.createCell(6).setCellValue(customer.getValidity());

                    EmailService.attachCount++;
                    currentAttachCount++;
                }
            }
            if (currentAttachCount > 0)
                myTrayIcon.displayMessage("Росгосстрах", "Прикреплено " + currentAttachCount + " пациентов", TrayIcon.MessageType.INFO);

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

    public void removeCustomersFromFile(List<RosGosStrahModel> customers) {
        for (RosGosStrahModel customer : customers) {
            removeCustomerFromFile("Росгосстрах", storageFileUrl, customer.getPolicyNumber(), 5);
        }
    }

}
