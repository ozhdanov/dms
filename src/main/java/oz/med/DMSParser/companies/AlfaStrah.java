package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.model.AlfaStrahModel;

import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
@Slf4j
public class AlfaStrah extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${alfastrah.storage}")
    private String storageFileUrl;
    @Value("${alfastrah.sender}")
    private String senderEmailTemplate;
    @Value("${alfastrah.liststemplate}")
    private String listsTemplate;
    @Value("${alfastrah.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${alfastrah.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${alfastrah.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${alfastrah.deattachfiletemplate}")
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

    public List<AlfaStrahModel> parseAttachListExcel(InputStream is) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        List<AlfaStrahModel> customers = new ArrayList<>();
        try {
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(is);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

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
                            row = rowIt.next();
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        break;
                    }

                    try {
                        log.info("Прикрепление пациента");
                        AlfaStrahModel alfaStrahModel = new AlfaStrahModel();
                        //Костыль против пробразования строки в число
                        Cell policyNumberCell = cellIterator.next();
                        policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                        String policyNumber = policyNumberCell.getStringCellValue();
                        alfaStrahModel.setPolicyNumber(policyNumber);

                        alfaStrahModel.setFio(cellIterator.next().toString());
                        alfaStrahModel.setDateOfBirth(format.parse(cellIterator.next().toString()));
                        alfaStrahModel.setAdress(cellIterator.next().toString());
                        alfaStrahModel.setPlaceOfWork(cellIterator.next().toString());
                        alfaStrahModel.setPolicyStartDate(format.parse(cellIterator.next().toString()));
                        alfaStrahModel.setPolicyEndDate(format.parse(cellIterator.next().toString()));
                        alfaStrahModel.setPolicyType(cellIterator.next().toString());
                        log.info(alfaStrahModel.toString());

                        customers.add(alfaStrahModel);
                    } catch (ParseException e) {
                        log.error("Ошибка парсинга строки", e);
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

    public List<AlfaStrahModel> parseDeattachListExcel(InputStream is) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        List<AlfaStrahModel> customers = new ArrayList<>();
        try {
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(is);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

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
                            row = rowIt.next();
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else if (cellIterator.hasNext()){

                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        break;
                    }

                    DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

                    try {
                        log.info("Открепление пациента");

                        AlfaStrahModel alfaStrahModel = new AlfaStrahModel();
                        //Костыль против пробразования строки в число
                        Cell policyNumberCell = cellIterator.next();
                        policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                        String policyNumber = policyNumberCell.getStringCellValue();
                        alfaStrahModel.setPolicyNumber(policyNumber);

                        alfaStrahModel.setFio(cellIterator.next().toString());
                        alfaStrahModel.setDateOfBirth(format.parse(cellIterator.next().toString()));
                        alfaStrahModel.setPlaceOfWork(cellIterator.next().toString());
                        alfaStrahModel.setDeattachDate(format.parse(cellIterator.next().toString()));
                        log.info(alfaStrahModel.toString());

                        customers.add(alfaStrahModel);
                    } catch (ParseException e) {
                        log.error("Ошибка парсинга строки", e);
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

    public void addCustomersToFile(List<AlfaStrahModel> customers) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(storageFileUrl));
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(inputStream);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            // we iterate on rows
            Iterator<Row> rowIt = sheet.iterator();

            boolean startOfDataFlag = false;

            while (rowIt.hasNext()) {
                //Пропускаем заголовок
                Row row = rowIt.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Cell cell;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        if (cell.toString().equals("№ полиса")) {
                            startOfDataFlag = true;
                            break;
                        }
                    }
                    continue;
                } else {
                    if(cellIterator.hasNext()) {

                        Cell policyNumberCell = cellIterator.next();

                        if(cellIterator.hasNext() && !policyNumberCell.toString().isEmpty()){
                            //Костыль против пробразования строки в число
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();
                            String fio = cellIterator.next().toString();

                            for (AlfaStrahModel customer : customers) {
                                if (policyNumber.equals(customer.getPolicyNumber())
                                        && fio.equals(customer.getFio()))
                                    customer.setNew(false);
                            }
                        }

                    }
                }
            }

            for (AlfaStrahModel customer : customers) {
                if (customer.isNew()) {
                    XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getFio());
                    row.createCell(2).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(3).setCellValue(customer.getAdress());
                    row.createCell(4).setCellValue(customer.getPlaceOfWork());
                    row.createCell(5).setCellValue(format.format(customer.getPolicyStartDate()));
                    row.createCell(6).setCellValue(format.format(customer.getPolicyEndDate()));
                    row.createCell(7).setCellValue(customer.getPolicyType());
                }
            }

            FileOutputStream outputStream = new FileOutputStream(storageFileUrl);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            workbook.close();
            inputStream.close();
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

    public void removeCustomersFromFile(List<AlfaStrahModel> customers) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileInputStream inputStream = null;
        List<Row> listOfRowsToRemove = new ArrayList<>();
        try {
            inputStream = new FileInputStream(new File(storageFileUrl));
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(inputStream);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            // we iterate on rows
            Iterator<Row> rowIt = sheet.iterator();

            boolean startOfDataFlag = false;

            while (rowIt.hasNext()) {
                //Пропускаем заголовок
                Row row = rowIt.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                Cell cell;
                if (!startOfDataFlag) {
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        if (cell.toString().equals("№ полиса")) {
                            startOfDataFlag = true;
                            break;
                        }
                    }
                    continue;
                } else {
                    if(cellIterator.hasNext()) {

                        Cell policyNumberCell = cellIterator.next();

                        if(cellIterator.hasNext() && !policyNumberCell.toString().isEmpty()) {
                            //Костыль против пробразования строки в число
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();

                            String fio = cellIterator.next().toString();

                            for (AlfaStrahModel customer : customers) {
                                if (policyNumber.equals(customer.getPolicyNumber())
                                        && fio.equals(customer.getFio()))
                                    listOfRowsToRemove.add(row);
                            }
                        }
                    }
                }
            }

            for(Row row: listOfRowsToRemove){
//                sheet.removeExcelRow(row);
                removeExcelRow(sheet, row.getRowNum());
            }

            FileOutputStream outputStream = new FileOutputStream(storageFileUrl);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            workbook.close();
            inputStream.close();
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

}