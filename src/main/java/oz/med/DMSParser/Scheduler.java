package oz.med.DMSParser;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import oz.med.DMSParser.services.EmailService;

@Component
public class Scheduler {

    @Autowired
    EmailService emailService;

//    @Scheduled(cron = "0 0/30 * * * ?")
    public void parsingEmailsJob(){

        emailService.handleEmails();
    }

}
