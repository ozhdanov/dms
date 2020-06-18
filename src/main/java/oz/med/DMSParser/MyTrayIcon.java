package oz.med.DMSParser;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.stereotype.Component;
import oz.med.DMSParser.services.EmailService;

import javax.annotation.PostConstruct;
import javax.swing.*;
import java.awt.*;
import java.net.URL;

@Component
public class MyTrayIcon extends TrayIcon {

    @Autowired
    EmailService emailService;

    private static final String IMAGE_PATH = "/Red-Cross.png";
    private static final String TOOLTIP = "Обработчик писем ДМС";

    public static PopupMenu popup;
    public static SystemTray tray;

    public MyTrayIcon(){
        super(createImage(IMAGE_PATH,TOOLTIP),TOOLTIP);
        popup = new PopupMenu();
        tray = SystemTray.getSystemTray();
    }

    @PostConstruct
    private void setup() throws AWTException{
        // popup.add(itemAbout);
        // here add the items to your popup menu. These extend MenuItem
        // popup.addSeparator();

        MenuItem startParsing = new MenuItem("Запустить обработку");
        popup.add(startParsing);
        startParsing.addActionListener(e -> {
                    this.setToolTip("Обработка...");
                    emailService.handleEmails();
                    this.setToolTip("Обработчик писем ДМС");
                }
        );

        MenuItem exitItem = new MenuItem("Выйти");
        popup.add(exitItem);
        exitItem.addActionListener(e -> {
            final int exitCode = 0;
            ExitCodeGenerator exitCodeGenerator = () -> exitCode;
            tray.remove(MyTrayIcon.this);
            System.exit(exitCode);

        });

        setPopupMenu(popup);
        tray.add(this);
    }

    protected static Image createImage(String path, String description){
        URL imageURL = MyTrayIcon.class.getResource(path);
        if(imageURL == null){
            System.err.println("Failed Creating Image. Resource not found: "+path);
            return null;
        }else {
            return new ImageIcon(imageURL,description).getImage();
        }
    }
}