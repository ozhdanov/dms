package oz.med.DMSParser;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;
import oz.med.DMSParser.services.EmailService;

@Component
public class Runner implements CommandLineRunner {
//
//    @Autowired
//    private JdbcTemplate jdbcTemplate;
    @Autowired
    EmailService emailService;

    @Override
    public void run(String... args) throws Exception {
//        System.setProperty("mail.mime.encodeparameters",  "false");
        System.setProperty("mail.mime.charset",  "utf-8");

//        log.info(testQuery());

//        testParsePDF();

        emailService.handleEmails();


    }

//    private List<String> testQuery(){
//        return jdbcTemplate
//                .queryForList("select name from customers;", String.class);
//    }

}
