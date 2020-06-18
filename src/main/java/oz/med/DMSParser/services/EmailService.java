package oz.med.DMSParser.services;

import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Service;
import oz.med.DMSParser.MyTrayIcon;
import oz.med.DMSParser.companies.*;
import oz.med.DMSParser.model.*;

import javax.mail.*;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeUtility;
import java.awt.*;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Properties;

@Service
@Slf4j
public class EmailService {

    @Value("${mailsListSize}")
    private Integer mailsListSize;

    private String saveDirectory = "D:\\temp\\mail";

    @Autowired
    MyTrayIcon myTrayIcon;

    @Autowired
    PdfParser pdfParser;

    @Autowired
    ExcelParser excelParser;

    @Autowired
    BestDoctor bestDoctor;
    @Autowired
    AlfaStrah alfaStrah;
    @Autowired
    RosGosStrah rosGosStrah;
    @Autowired
    InGosStrah inGosStrah;
    @Autowired
    Absolut absolut;
    @Autowired
    Sogaz sogaz;
    @Autowired
    Reso reso;
    @Autowired
    Soglasie soglasie;

    @Value("${spring.mail.protocol}")
    private String protocol;
    @Value("${spring.mail.host}")
    private String host;
    @Value("${spring.mail.port}")
    private String port;
    @Value("${spring.mail.username}")
    private String userName;
    @Value("${spring.mail.password}")
    private String password;

    public static int attachCount = 0;
    public static int deattachCount = 0;

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
    public void handleEmails() {

        log.info("Начало обработки писем");

        attachCount = 0;
        deattachCount = 0;

        Properties properties = getServerProperties(protocol, host, port);
        Session session = Session.getDefaultInstance(properties);

        Message[] messages;

        try {
            // connects to the message store
            Store store = session.getStore(protocol);
            store.connect(userName, password);

            // opens the inbox folder
            Folder folderInbox = store.getFolder("INBOX");
            folderInbox.open(Folder.READ_ONLY);

            int messageCount = folderInbox.getMessageCount();

            messages = folderInbox.getMessages(messageCount - mailsListSize, messageCount);

            log.info("Обработка {} последних писем", mailsListSize);

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

                if (!(
//                        bestDoctor.isListsMail(from, subject) ||
//                                alfaStrah.isListsMail(from, subject) ||
//                                rosGosStrah.isListsMail(from, subject) ||
//                                inGosStrah.isListsMail(from, subject) ||
//                                absolut.isListsMail(from, subject) ||
//                                sogaz.isListsMail(from, subject) ||
//                                reso.isListsMail(from, subject) ||
                                soglasie.isListsMail(from, subject)
                )) continue;

//                log.info("\t From: " + from);
//                log.info("\t Subject: " + subject);

                // store attachment file name, separated by comma
                String attachFiles = "";

                if (contentType.contains("multipart")) {
                    // content may contain attachments
                    Multipart multiPart;
                    try {
                        multiPart = (Multipart) message.getContent();
                    } catch (IOException e){
                        log.error("Ошибка чтения письма", e);
                        continue;
                    }
                    int numberOfParts = multiPart.getCount();
                    for (int partCount = 0; partCount < numberOfParts; partCount++) {
                        MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
                        if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())
                                || Part.INLINE.equalsIgnoreCase(part.getDisposition())
                                || (part.getDisposition() == null && part.getFileName() != null)) {
                            String fileName = "";

                            try {
                                //У всех файл прикреплен по разному и разные кодировки
                                if ((inGosStrah.isListsMail(from, subject) || absolut.isListsMail(from, subject) || sogaz.isListsMail(from, subject) || soglasie.isListsMail(from, subject))
                                        && Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition()))
                                    fileName = MimeUtility.decodeText(part.getFileName());
                                else if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition()))
                                    fileName = new String(part.getFileName().getBytes("ISO-8859-1"));
                                else if (Part.INLINE.equalsIgnoreCase(part.getDisposition()) || (part.getDisposition() == null && part.getFileName() != null))
                                    fileName = MimeUtility.decodeText(part.getFileName());
                            } catch (UnsupportedEncodingException e){
                                log.error("Ошибка обработки имени файла", e);
                                continue;
                            }

                            attachFiles += fileName + ", ";

                            //БэстДоктор
                            try {
                                if (bestDoctor.isListsMail(from, subject)) {
                                    if (bestDoctor.isAttachFile(fileName)) {
                                        List<BestDoctorModel> bestDoctorModels = bestDoctor.parseAttachListExcel(part.getInputStream());
                                        bestDoctor.addCustomersToFile(bestDoctorModels);
                                    } else if (bestDoctor.isDeattachFile(fileName)) {
                                        List<BestDoctorModel> bestDoctorModels = bestDoctor.parseDeattachListExcel(part.getInputStream());
                                        if (bestDoctorModels.size() > 0)
                                            bestDoctor.removeCustomersFromFile(bestDoctorModels);
                                    }
                                }
                                //Альфа
                                else if (alfaStrah.isAttachListMail(from, subject)) {
                                    if (alfaStrah.isAttachFile(fileName)) {
                                        List<AlfaStrahModel> alfaStrahModels = alfaStrah.parseAttachListExcel(part.getInputStream());
                                        alfaStrah.addCustomersToFile(alfaStrahModels);
                                    }
                                } else if (alfaStrah.isDeattachListMail(from, subject)) {
                                    if (alfaStrah.isDeattachFile(fileName)) {
                                        List<AlfaStrahModel> alfaStrahModels = alfaStrah.parseDeattachListExcel(part.getInputStream());
                                        if (alfaStrahModels.size() > 0)
                                            alfaStrah.removeCustomersFromFile(alfaStrahModels);
                                    }
                                }
                                //РосГосСтрах
                                else if (rosGosStrah.isListsMail(from, subject)) {
                                    if (rosGosStrah.isAttachFile(fileName)) {
                                        List<RosGosStrahModel> rosGosStrahModels = rosGosStrah.parseAttachListExcel(part.getInputStream());
                                        rosGosStrah.addCustomersToFile(rosGosStrahModels);
                                    } else if (rosGosStrah.isDeattachFile(fileName)) {
                                        List<RosGosStrahModel> rosGosStrahModels = rosGosStrah.parseDeattachListExcel(part.getInputStream());
                                        if (rosGosStrahModels.size() > 0)
                                            rosGosStrah.removeCustomersFromFile(rosGosStrahModels);
                                    }
                                }
                                //ИнГосСтрах
                                else if (inGosStrah.isAttachListMail(from, subject)) {
                                    if (inGosStrah.isAttachFile(fileName)) {
                                        List<InGosStrahModel> inGosStrahModels = inGosStrah.parseAttachListExcel(part.getInputStream());
                                        inGosStrah.addCustomersToFile(inGosStrahModels);
                                    }
                                } else if (inGosStrah.isDeattachListMail(from, subject)) {
                                    if (inGosStrah.isDeattachFile(fileName)) {
                                        List<InGosStrahModel> inGosStrahModels = inGosStrah.parseDeattachListExcel(part.getInputStream());
                                        if (inGosStrahModels.size() > 0)
                                            inGosStrah.removeCustomersFromFile(inGosStrahModels);
                                    }
                                }
                                //Абсолют
                                else if (absolut.isAttachListMail(from, subject)) {
                                    if (absolut.isAttachFile(fileName)) {
                                        List<AbsolutModel> absolutModels = absolut.parseAttachListExcel(part.getInputStream());
                                        absolut.addCustomersToFile(absolutModels);
                                    }
                                } else if (absolut.isDeattachListMail(from, subject)) {
                                    if (absolut.isDeattachFile(fileName)) {
                                        List<AbsolutModel> absolutModels = absolut.parseDeattachListExcel(part.getInputStream());
                                        if (absolutModels.size() > 0) absolut.removeCustomersFromFile(absolutModels);
                                    }
                                }
                                //Согаз
                                else if (sogaz.isAttachListMail(from, subject)) {
                                    if (sogaz.isAttachFile(fileName)) {
                                        List<SogazModel> sogazModels = sogaz.parseAttachListExcel(part.getInputStream());
                                        sogaz.addCustomersToFile(sogazModels);
                                    }
                                } else if (sogaz.isDeattachListMail(from, subject)) {
                                    if (sogaz.isDeattachFile(fileName)) {
                                        List<SogazModel> sogazModels = sogaz.parseDeattachListExcel(part.getInputStream());
                                        if (sogazModels.size() > 0) sogaz.removeCustomersFromFile(sogazModels);
                                    }
                                }
                                //Ресо
                                else if (reso.isAttachListMail(from, subject)) {
                                    if (reso.isAttachFile(fileName)) {
                                        List<ResoModel> resoModels = reso.parseAttachListRTF(part.getInputStream());
                                        reso.addCustomersToFile(resoModels);
                                    }
                                } else if (reso.isDeattachListMail(from, subject)) {
                                    if (reso.isDeattachFile(fileName)) {
                                        List<ResoModel> resoModels = reso.parseDeattachListExcel(part.getInputStream());
                                        if (resoModels.size() > 0) reso.removeCustomersFromFile(resoModels);
                                    }
                                }
                                //Согласие
                                else if (soglasie.isListsMail(from, subject)) {
                                    if (soglasie.isAttachFile(fileName)) {
                                        List<SoglasieModel> soglasieModels = soglasie.parseAttachListExcel(part.getInputStream());
                                        soglasie.addCustomersToFile(soglasieModels);
                                    }
//                                    if (soglasie.isDeattachFile(fileName)) {
//                                        List<SoglasieModel> soglasieModels = soglasie.parseDeattachListExcel(part.getInputStream());
//                                        if (soglasieModels.size() > 0) soglasie.removeCustomersFromFile(soglasieModels);
//                                    }
                                }
                            } catch (IOException e) {
                                log.error("Ошибка извлечения файла от " + from, e);
                            }

                            //todo
//                            part.saveFile(saveDirectory + File.separator
//                                    + new DateTime(message.getSentDate()).toString("yyyy-MM-dd_HH-mm-ss") + "_" + fileName);

                        }
//                        else if (part.getContent() != null){
//                            // this part may be the message content
//                            messageContent = part.getContent().toString();
//                        }
                    }

                    if (attachFiles.length() > 1) {
                        attachFiles = attachFiles.substring(0, attachFiles.length() - 2);
                    }
                }
//                else if (contentType.contains("text/plain")
//                        || contentType.contains("text/html")) {
//                    Object content = message.getContent();
//                    if (content != null) {
//                        messageContent = content.toString();
//                    }
//                }

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

            if (attachCount > 0)
                myTrayIcon.displayMessage("ДМС", "Прикреплено " + attachCount + " пациентов", TrayIcon.MessageType.INFO);
            if (deattachCount > 0)
                myTrayIcon.displayMessage("ДМС", "Откреплено " + deattachCount + " пациентов", TrayIcon.MessageType.INFO);

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
