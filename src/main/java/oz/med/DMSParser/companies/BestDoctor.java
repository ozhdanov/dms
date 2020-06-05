package oz.med.DMSParser.companies;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.model.BestDoctorModel;

import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class BestDoctor {

    @Value("${storage.bestdoctor}")
    private String storageFileUrl;

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    public boolean isAttachMail(String from, String subject) {
        if (from.contains("list@bestdoctor.ru") && subject.toUpperCase().contains("СПИСКИ"))
            return true;
        else
            return false;
    }

    public List<BestDoctorModel> parseAttachListExcel(InputStream is) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        List<BestDoctorModel> customers = new ArrayList<>();
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
                        if (cell.toString().equals("№п/п")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else {

                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        break;
                    }

                    DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

                    try {
                        System.out.println("Прикрепление пациента");
                        BestDoctorModel bestDoctorModel = new BestDoctorModel();
                        bestDoctorModel.setPolicyNumber(String.valueOf((int) cellIterator.next().getNumericCellValue()));
                        bestDoctorModel.setSurname(cellIterator.next().toString());
                        bestDoctorModel.setName(cellIterator.next().toString());
                        bestDoctorModel.setPatronymic(cellIterator.next().toString());
                        bestDoctorModel.setSex(cellIterator.next().toString());
                        bestDoctorModel.setDateOfBirth(format.parse(cellIterator.next().toString()));
                        bestDoctorModel.setAdress(cellIterator.next().toString());
                        bestDoctorModel.setPhoneNumber(cellIterator.next().toString());
                        bestDoctorModel.setPolicyStartDate(format.parse(cellIterator.next().toString()));
                        bestDoctorModel.setPolicyEndDate(format.parse(cellIterator.next().toString()));
                        bestDoctorModel.setPlaceOfWork(cellIterator.next().toString());
                        System.out.println(bestDoctorModel.toString());

                        customers.add(bestDoctorModel);
                    } catch (ParseException e) {
                        System.out.println(e);
                    }

                }

                prewRowIndex = row.getRowNum();

            }

            System.out.println();

            workbook.close();
            is.close();
        } catch (IOException e) {
            System.out.println("Не удалось распарсить документ");
            System.out.println(e);
        } finally {
            try {
                workbook.close();
                is.close();
            } catch (Exception e) {
            }
        }

        return customers;

    }

    public List<BestDoctorModel> parseDeattachListExcel(InputStream is) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        List<BestDoctorModel> customers = new ArrayList<>();
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
                        if (cell.toString().equals("№п/п")) {
                            startOfDataFlag = true;
                            prewRowIndex = row.getRowNum();
                            break;
                        }
                    }
                    continue;
                } else {

                    cell = cellIterator.next();

                    if (cell.getRowIndex() - prewRowIndex > 1) {
                        startOfDataFlag = false;
                        break;
                    }

                    DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

                    try {
                        System.out.println("Открепление пациента");
                        BestDoctorModel bestDoctorModel = new BestDoctorModel();
                        bestDoctorModel.setPolicyNumber(cellIterator.next().toString());
                        bestDoctorModel.setSurname(cellIterator.next().toString());
                        bestDoctorModel.setName(cellIterator.next().toString());
                        bestDoctorModel.setPatronymic(cellIterator.next().toString());
                        bestDoctorModel.setSex(cellIterator.next().toString());
                        bestDoctorModel.setDateOfBirth(format.parse(cellIterator.next().toString()));
                        bestDoctorModel.setPolicyEndDate(format.parse(cellIterator.next().toString()));
                        System.out.println(bestDoctorModel.toString());

                        customers.add(bestDoctorModel);
                    } catch (ParseException e) {
                        System.out.println(e);
                    }

                }
                prewRowIndex = row.getRowNum();
            }
            System.out.println();

            workbook.close();
            is.close();
        } catch (IOException e) {
            System.out.println("Не удалось распарсить документ");
            System.out.println(e);
        } finally {
            try {
                workbook.close();
                is.close();
            } catch (Exception e) {
            }
        }

        return customers;

    }

    public void addCustomersToFile(List<BestDoctorModel> customers) {

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
                        String policyNumber = cellIterator.next().toString();
                        String surname = cellIterator.next().toString();
                        String name = cellIterator.next().toString();

                        for (BestDoctorModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber())
                                    && surname.equals(customer.getSurname())
                                    && name.equals(customer.getName()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            for (BestDoctorModel customer : customers) {
                if (customer.isNew()) {
                    XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getSurname());
                    row.createCell(2).setCellValue(customer.getName());
                    row.createCell(3).setCellValue(customer.getPatronymic());
                    row.createCell(4).setCellValue(customer.getSex());
                    row.createCell(5).setCellValue(format.format(customer.getDateOfBirth()));
                    row.createCell(6).setCellValue(customer.getAdress());
                    row.createCell(7).setCellValue(customer.getPhoneNumber());
                    row.createCell(8).setCellValue(format.format(customer.getPolicyStartDate()));
                    row.createCell(9).setCellValue(format.format(customer.getPolicyEndDate()));
                    row.createCell(10).setCellValue(customer.getPlaceOfWork());
                }
            }

            FileOutputStream outputStream = new FileOutputStream(storageFileUrl);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println();

            workbook.close();
            inputStream.close();
        } catch (IOException e) {
            System.out.println("Не удалось распарсить документ");
            System.out.println(e);
        } finally {
            try {
                workbook.close();
                inputStream.close();
            } catch (Exception e) {
            }
        }

    }

    public void removeCustomersFromFile(List<BestDoctorModel> customers) {

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
                        String policyNumber = cellIterator.next().toString();
                        String surname = cellIterator.next().toString();
                        String name = cellIterator.next().toString();

                        for (BestDoctorModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber())
                                    && surname.equals(customer.getSurname())
                                    && name.equals(customer.getName()))
                                listOfRowsToRemove.add(row);
                        }
                    }
                }
            }

            for(Row row: listOfRowsToRemove){
//                sheet.removeRow(row);
                removeRow(sheet, row.getRowNum());
            }

            FileOutputStream outputStream = new FileOutputStream(storageFileUrl);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println();

            workbook.close();
            inputStream.close();
        } catch (IOException e) {
            System.out.println("Не удалось распарсить документ");
            System.out.println(e);
        } finally {
            try {
                workbook.close();
                inputStream.close();
            } catch (Exception e) {
            }
        }
    }

    public void removeRow(Sheet sheet, int rowIndex) {
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
