package oz.med.DMSParser.companies;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.sax.BodyContentHandler;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.xml.sax.SAXException;
import oz.med.DMSParser.model.ResoModel;
import oz.med.DMSParser.services.EmailService;
import oz.med.DMSParser.services.MyTrayIcon;
import oz.med.DMSParser.services.RtfParser;

import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.Document;
import javax.swing.text.rtf.RTFEditorKit;
import java.awt.*;
import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
@Slf4j
public class Reso extends Company {

    private DateFormat format = new SimpleDateFormat("dd.MM.yyyy");

    @Value("${reso.storage}")
    private String storageFileUrl;
    @Value("${reso.sender}")
    private String senderEmailTemplate;
    @Value("${reso.liststemplate}")
    private String listsTemplate;
    @Value("${reso.attachlisttemplate}")
    private String attachListTemplate;
    @Value("${reso.deattachlisttemplate}")
    private String deattachListTemplate;
    @Value("${reso.attachfiletemplate}")
    private String attachFileTemplate;
    @Value("${reso.deattachfiletemplate}")
    private String deattachFileTemplate;

    @Autowired
    RtfParser rtfParser;

    @Autowired
    MyTrayIcon myTrayIcon;

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

    public List<ResoModel> parseAttachListRTF(InputStream is) {
        List<ResoModel> customers = new ArrayList<>();
        try {

            String text = rtfParser.parseToPlainText(is);

            String placeOfWork = text.substring(
                    text.indexOf("застрахованных от") + "застрахованных от".length() + 1,
                    text.indexOf("(договор со страхователем") - 1 );

            String policyType = text.substring(
                    text.indexOf("Страховая программа.") + "Страховая программа.".length() + 1);
            policyType = policyType.substring(0, policyType.indexOf("\n"));


            while (text.contains("Период страхования:")){
                text = text.substring(text.indexOf("Период страхования:") + "Период страхования:".length() + 1);
                String validity = text.substring(0, text.indexOf("\n"));

                String customersStr = text.substring(text.indexOf("Телефон\n") + "Телефон\n".length(), text.indexOf("\n\n\n") + "\n\n\n".length());

                while (customersStr.contains("\n\n")){
                    String customerStr = customersStr.substring(0, customersStr.indexOf("\n\n") + 1);

                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String policyNumber = customerStr.substring(0, customerStr.indexOf("\n"));
                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String fio = customerStr.substring(0, customerStr.indexOf("\n"));
                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String dateOfBirth = customerStr.substring(0, customerStr.indexOf("\n"));
                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String adress = customerStr.substring(0, customerStr.indexOf("\n"));

                    ResoModel resoModel = new ResoModel();
                    resoModel.setFio(fio);
                    resoModel.setPolicyNumber(policyNumber);
                    resoModel.setDateOfBirth(dateOfBirth);
                    resoModel.setAdress(adress);
                    resoModel.setPlaceOfWork(placeOfWork);
                    resoModel.setValidity(validity);
                    resoModel.setPolicyType(policyType);
                    customers.add(resoModel);

                    if(customersStr.contains("\n\n"))
                        customersStr = customersStr.substring(customersStr.indexOf("\n\n") + "\n\n".length(), customersStr.length());
                }

            }
        } catch (Exception e){
            log.error("Не удалось распарсить документ", e);
        }

        return customers;
    }

    public List<ResoModel> parseDeattachListExcel(InputStream is) {
        List<ResoModel> customers = new ArrayList<>();
        try {
            String text = rtfParser.parseToPlainText(is);
            while (text.contains("Дата открепления")){
                text = text.substring(text.indexOf("Дата открепления") + "Дата открепления".length() + 1);

                String customersStr = text.substring(text.indexOf("Дата рождения\n") + "Дата рождения\n".length(), text.indexOf("\n\n\n") + "\n\n\n".length());

                while (customersStr.contains("\n\n")){
                    String customerStr = customersStr.substring(0, customersStr.indexOf("\n\n") + 1);

                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String policyNumber = customerStr.substring(0, customerStr.indexOf("\n"));
                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String fio = customerStr.substring(0, customerStr.indexOf("\n"));
                    customerStr = customerStr.substring(customerStr.indexOf("\n") + "\n".length(), customerStr.length());
                    String dateOfBirth = customerStr.substring(0, customerStr.indexOf("\n"));

                    ResoModel resoModel = new ResoModel();
                    resoModel.setFio(fio);
                    resoModel.setPolicyNumber(policyNumber);
                    resoModel.setDateOfBirth(dateOfBirth);
                    customers.add(resoModel);

                    if(customersStr.contains("\n\n"))
                        customersStr = customersStr.substring(customersStr.indexOf("\n\n") + "\n\n".length(), customersStr.length());
                }

            }
        } catch (Exception e){
            log.error("Не удалось распарсить документ", e);
        }
        return customers;
    }

    public void addCustomersToFile(List<ResoModel> customers) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(storageFileUrl));
            // we create an XSSF Workbook object for our XLSX Excel File
            workbook = new XSSFWorkbook(inputStream);
            // we get first sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() > 0 && !isRowEmpty(row)) {
                    Cell policyNumberCell = row.getCell(0);
                    policyNumberCell.setCellType(CellType.STRING);
                    String policyNumber = policyNumberCell.getStringCellValue();
                    if(!policyNumber.toString().isEmpty()) {
                        for (ResoModel customer : customers) {
                            if (policyNumber.equals(customer.getPolicyNumber()))
                                customer.setNew(false);
                        }
                    }
                }
            }

            for (ResoModel customer : customers) {
                if (customer.isNew()) {
                    XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    row.createCell(0).setCellValue(customer.getPolicyNumber());
                    row.createCell(1).setCellValue(customer.getFio());
                    row.createCell(2).setCellValue(customer.getDateOfBirth());
                    row.createCell(3).setCellValue(customer.getAdress());
                    row.createCell(4).setCellValue(customer.getPlaceOfWork());
                    row.createCell(5).setCellValue(customer.getValidity());
                    row.createCell(6).setCellValue(customer.getPolicyType());

                    EmailService.attachCount++;
                }
            }

            FileOutputStream outputStream = new FileOutputStream(storageFileUrl);
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

    public void removeCustomersFromFile(List<ResoModel> customers) {
        for (ResoModel customer : customers) {
            removeCustomerFromFile(storageFileUrl, customer.getPolicyNumber(), 0);
        }
    }

}
