package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.model.RosGosStrahModel;

import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
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
                            log.info("Прикрепление пациента");
                            RosGosStrahModel rosGosStrahModel = new RosGosStrahModel();

                            rosGosStrahModel.setFio(cellIterator.next().toString());
                            rosGosStrahModel.setSex(cellIterator.next().toString());
                            rosGosStrahModel.setDateOfBirth(cellIterator.next().getDateCellValue());
                            rosGosStrahModel.setAdress(cellIterator.next().toString());
                            rosGosStrahModel.setPhoneNumber(cellIterator.next().toString());

                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();
                            rosGosStrahModel.setPolicyNumber(policyNumber);

                            log.info(rosGosStrahModel.toString());

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
                            log.info("Открепление пациента");

                            RosGosStrahModel rosGosStrahModel = new RosGosStrahModel();

                            rosGosStrahModel.setFio(cellIterator.next().toString());
                            rosGosStrahModel.setSex(cellIterator.next().toString());
                            rosGosStrahModel.setDateOfBirth(cellIterator.next().getDateCellValue());
                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();
                            rosGosStrahModel.setPolicyNumber(policyNumber);

                            log.info(rosGosStrahModel.toString());

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
                        if (cell.toString().equals("ФИО")) {
                            startOfDataFlag = true;
                            break;
                        }
                    }
                    continue;
                } else {
                    if(cellIterator.hasNext()) {

                        Cell fioCell = cellIterator.next();
                        String fio = fioCell.getStringCellValue();

                        if(cellIterator.hasNext() && !fioCell.toString().isEmpty()) {

                            cellIterator.next();
                            cellIterator.next();
                            cellIterator.next();
                            cellIterator.next();

                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();

                            for (RosGosStrahModel customer : customers) {
                                if (policyNumber.equals(customer.getPolicyNumber())
                                        && fio.equals(customer.getFio()))
                                    customer.setNew(false);
                            }
                        }

                    }
                }
            }

            for (RosGosStrahModel customer : customers) {
                if (customer.isNew()) {
                    XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    row.createCell(0).setCellValue(customer.getFio());
                    row.createCell(1).setCellValue(customer.getSex());
                    if(customer.getDateOfBirth() != null) row.createCell(2).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(3).setCellValue(customer.getAdress());
                    row.createCell(4).setCellValue(customer.getPhoneNumber());
                    row.createCell(5).setCellValue(customer.getPolicyNumber());
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

    public void removeCustomersFromFile(List<RosGosStrahModel> customers) {

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
                        if (cell.toString().equals("ФИО")) {
                            startOfDataFlag = true;
                            break;
                        }
                    }
                    continue;
                } else {
                    if(cellIterator.hasNext()) {

                        Cell fioCell = cellIterator.next();
                        String fio = fioCell.getStringCellValue();

                        if(cellIterator.hasNext() && !fioCell.toString().isEmpty()) {

                            cellIterator.next();
                            cellIterator.next();
                            cellIterator.next();
                            cellIterator.next();

                            //Костыль против пробразования строки в число
                            Cell policyNumberCell = cellIterator.next();
                            policyNumberCell.setCellType(Cell.CELL_TYPE_STRING);
                            String policyNumber = policyNumberCell.getStringCellValue();

                            for (RosGosStrahModel customer : customers) {
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
