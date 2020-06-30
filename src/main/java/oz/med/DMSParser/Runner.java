package oz.med.DMSParser;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;
import oz.med.DMSParser.companies.InGosStrah;
import oz.med.DMSParser.model.InGosStrahModel;
import oz.med.DMSParser.services.EmailService;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

@Component
public class Runner implements CommandLineRunner {
//
//    @Autowired
//    private JdbcTemplate jdbcTemplate;
    @Autowired
    EmailService emailService;
    @Autowired
    InGosStrah inGosStrah;

    @Override
    public void run(String... args) throws Exception {
//        System.setProperty("mail.mime.encodeparameters",  "false");
        System.setProperty("mail.mime.charset",  "utf-8");

//        log.info(testQuery());

//        testParsePDF();

        emailService.handleEmails();

//        parseIngosstrahFile();


    }

    private void parseIngosstrahFile() throws FileNotFoundException {
        File initialFile = new File("D:\\temp\\mail\\src\\Ингосстрах списки.XLS");
        InputStream targetStream = new FileInputStream(initialFile);
        List<InGosStrahModel> inGosStrahModels = inGosStrah.parseAttachListExcel(targetStream);
        inGosStrah.addCustomersToFile(inGosStrahModels);
    }

//    private List<String> testQuery(){
//        return jdbcTemplate
//                .queryForList("select name from customers;", String.class);
//    }

}
