package oz.med.DMSParser.services;

import lombok.extern.slf4j.Slf4j;
import org.joda.time.DateTime;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.companies.BestDoctor;
import oz.med.DMSParser.model.BestDoctorModel;

import javax.mail.*;
import javax.mail.internet.MimeBodyPart;
import java.io.File;
import java.io.IOException;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Properties;

@Service
@Slf4j
public class EmailService {

    private Format formatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");

    private String saveDirectory = "D:\\temp\\mail";

    @Autowired
    PdfParser pdfParser;

    @Autowired
    ExcelParser excelParser;

    @Autowired
    BestDoctor bestDoctor;

    /**
     * Returns a Properties object which is configured for a POP3/IMAP server
     *
     * @param protocol either "imap" or "pop3"
     * @param host
     * @param port
     * @return a Properties object
     */
    private Properties getServerProperties(String protocol, String host,
                                           String port) {
        Properties properties = new Properties();

        // server setting
        properties.put(String.format("mail.%s.host", protocol), host);
        properties.put(String.format("mail.%s.port", protocol), port);

        // SSL setting
        properties.setProperty(
                String.format("mail.%s.socketFactory.class", protocol),
                "javax.net.ssl.SSLSocketFactory");
        properties.setProperty(
                String.format("mail.%s.socketFactory.fallback", protocol),
                "false");
        properties.setProperty(
                String.format("mail.%s.socketFactory.port", protocol),
                String.valueOf(port));

        return properties;
    }

    /**
     * Downloads new messages and fetches details for each message.

     */
    public void handleEmails() throws IOException {

        log.info("Начало обработки писем");

        // for IMAP
        String protocol = "imap";
        String host = "imap.yandex.ru";
        String port = "993";


        String userName = "info@denttime.ru";
        String password = "23Ja8Uq(";

        Properties properties = getServerProperties(protocol, host, port);
        Session session = Session.getDefaultInstance(properties);

        Message[] messages = {};

        try {
            // connects to the message store
            Store store = session.getStore(protocol);
            store.connect(userName, password);

            // opens the inbox folder
            Folder folderInbox = store.getFolder("INBOX");
            folderInbox.open(Folder.READ_ONLY);

            int messageCount = folderInbox.getMessageCount();

            messages = folderInbox.getMessages(messageCount - 200, messageCount);

            log.info("Обработка 300 последних писем");

            for (int i = 0; i < messages.length; i++) {
                System.out.print(".");
                if (i%50 == 0) System.out.println("");
                Message message = messages[i];
                Address[] fromAddress = message.getFrom();
                String from = fromAddress[0].toString();

                String subject = message.getSubject();
                String toList = parseAddresses(message
                        .getRecipients(Message.RecipientType.TO));
                String ccList = parseAddresses(message
                        .getRecipients(Message.RecipientType.CC));
                String sentDate = message.getSentDate().toString();

                String contentType = message.getContentType();
                String messageContent = "";

                if (!(bestDoctor.isAttachMail(from, subject) ||
                        bestDoctor.isAttachMail(from, subject)
                )) continue;

                log.info("\t From: " + from);
                log.info("\t Subject: " + subject);

                // store attachment file name, separated by comma
                String attachFiles = "";

                if (contentType.contains("multipart")) {
                    // content may contain attachments
                    Multipart multiPart = (Multipart) message.getContent();
                    int numberOfParts = multiPart.getCount();
                    for (int partCount = 0; partCount < numberOfParts; partCount++) {
                        MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
                        if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
                            // this part is attachment
                            String fileName = new String(part.getFileName().getBytes("ISO-8859-1"));
                            log.info(fileName);
                            attachFiles += fileName + ", ";

                            if (bestDoctor.isAttachMail(from, subject)){
                                if(fileName.toUpperCase().contains("ПРИКРЕПЛЕНИЕ")) {
                                    List<BestDoctorModel> bestDoctorModels = bestDoctor.parseAttachListExcel(part.getInputStream());
                                    bestDoctor.addCustomersToFile(bestDoctorModels);
                                }
                                else if (fileName.toUpperCase().contains("ОТКРЕПЛЕНИЕ")) {
                                    List<BestDoctorModel> bestDoctorModels = bestDoctor.parseDeattachListExcel(part.getInputStream());
                                    bestDoctor.removeCustomersFromFile(bestDoctorModels);
                                }
                            }


                            part.saveFile(saveDirectory + File.separator
                                    + new DateTime(message.getSentDate()).toString("yyyy-MM-dd_HH-mm-ss") + "_" + fileName);

                        } else {
                            // this part may be the message content
                            messageContent = part.getContent().toString();
                        }
                    }

                    if (attachFiles.length() > 1) {
                        attachFiles = attachFiles.substring(0, attachFiles.length() - 2);
                    }
                } else if (contentType.contains("text/plain")
                        || contentType.contains("text/html")) {
                    Object content = message.getContent();
                    if (content != null) {
                        messageContent = content.toString();
                    }
                }

//                // print out details of each message
//                log.info();
//                log.info("Message #" + (i + 1) + ":");
//                log.info("\t From: " + from);
//                log.info("\t To: " + toList);
//                log.info("\t CC: " + ccList);
//                log.info("\t Subject: " + subject);
//                log.info("\t Sent Date: " + sentDate);
//                log.info();
//                log.info("\t Message: " + messageContent);
//                log.info("\t Attachments: " + attachFiles);
            }

            // disconnect
            folderInbox.close(false);
            store.close();

            log.info("Окончание обработки писем");

        } catch (NoSuchProviderException ex) {
            log.error("No provider for protocol: " + protocol, ex);
        } catch (MessagingException ex) {
            log.error("Could not connect to the message store", ex);
        }

    }

    /**
     * Returns a list of addresses in String format separated by comma
     *
     * @param address an array of Address objects
     * @return a string represents a list of addresses
     */
    private String parseAddresses(Address[] address) {
        String listAddress = "";

        if (address != null) {
            for (int i = 0; i < address.length; i++) {
                listAddress += address[i].toString() + ", ";
            }
        }
        if (listAddress.length() > 1) {
            listAddress = listAddress.substring(0, listAddress.length() - 2);
        }

        return listAddress;
    }


}
