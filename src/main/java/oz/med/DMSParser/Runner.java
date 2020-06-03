package oz.med.DMSParser;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.Properties;

@Component
public class Runner implements CommandLineRunner {

    @Autowired
    private JdbcTemplate jdbcTemplate;

    @Override
    public void run(String... args) throws Exception {

//        System.out.println(testQuery());

//        testParsePDF();


        // for IMAP
        String protocol = "imap";
        String host = "imap.yandex.ru";
        String port = "993";


        String userName = "olegzhdanov1@yandex.ru";
        String password = "vombat51";

        String saveDirectory = "D:\\temp\\mail";

        EmailReceiver receiver = new EmailReceiver();
        receiver.setSaveDirectory(saveDirectory);
        receiver.downloadEmails(protocol, host, port, userName, password);

    }

    private List<String> testQuery(){
        return jdbcTemplate
                .queryForList("select name from customers;", String.class);
    }

    public void testParsePDF(){
        PDDocument pdDoc = null;
        COSDocument cosDoc = null;
        PDFTextStripper pdfStripper;
        String parsedText;
        File folder = new File("D:\\temp\\mail");

        for (File fileEntry : folder.listFiles()) {
            if (!fileEntry.isDirectory()) {
                System.out.println(fileEntry.getName());

                try {
                    pdDoc = PDDocument.load(fileEntry);
                    pdfStripper = new PDFTextStripper();

                    parsedText = pdfStripper.getText(pdDoc);
//            System.out.println(parsedText.replaceAll("[^A-Za-z0-9. ]+", ""));
//            System.out.println(parsedText);

                    String fio = "";

                    String rawF = parsedText.substring(parsedText.indexOf("Застрахованный") + "Застрахованный".length(),
                            parsedText.indexOf("фамилия дата рождения"));
                    fio += rawF.replaceAll("\r|\n|[0-9]|[.!?]|[ ]", "");

                    String rawI = parsedText.substring(parsedText.indexOf("фамилия дата рождения") + "фамилия дата рождения".length(),
                            parsedText.indexOf("имя категория VIP"));
                    rawI = rawI.substring(0, rawI.indexOf( " "));
                    fio += " " + rawI.replaceAll("\r|\n|[0-9]|[.!?\\-]|[ ]", "");

                    String rawO = parsedText.substring(parsedText.indexOf("имя категория VIP") + "имя категория VIP".length(),
                            parsedText.indexOf("отчество номер полиса ДМС"));
                    rawO = rawO.substring(0, rawO.indexOf( " "));
                    fio += " " + rawO.replaceAll("\r|\n|[0-9]|[.!?\\-]|[ ]", "");

                    System.out.println(fio);
                } catch (Exception e) {
                    e.printStackTrace();
                    try {
                        if (cosDoc != null)
                            cosDoc.close();
                        if (pdDoc != null)
                            pdDoc.close();
                    } catch (Exception e1) {
                        e1.printStackTrace();
                    }

                }
            }
        }


    }
}
