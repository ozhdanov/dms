package oz.med.DMSParser.services;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;

@Service
public class PdfParser {

    public String getTextFromPDF(InputStream is) {
        PDDocument pdDoc = null;
        COSDocument cosDoc = null;
        PDFTextStripper pdfStripper;
        String parsedText;

        try {
            pdDoc = PDDocument.load(is);
            pdfStripper = new PDFTextStripper();

            parsedText = pdfStripper.getText(pdDoc);


////            System.out.println(parsedText.replaceAll("[^A-Za-z0-9. ]+", ""));
////            System.out.println(parsedText);
//
//                    String fio = "";
//
//                    String rawF = parsedText.substring(parsedText.indexOf("Застрахованный") + "Застрахованный".length(),
//                            parsedText.indexOf("фамилия дата рождения"));
//                    fio += rawF.replaceAll("\r|\n|[0-9]|[.!?]|[ ]", "");
//
//                    String rawI = parsedText.substring(parsedText.indexOf("фамилия дата рождения") + "фамилия дата рождения".length(),
//                            parsedText.indexOf("имя категория VIP"));
//                    rawI = rawI.substring(0, rawI.indexOf( " "));
//                    fio += " " + rawI.replaceAll("\r|\n|[0-9]|[.!?\\-]|[ ]", "");
//
//                    String rawO = parsedText.substring(parsedText.indexOf("имя категория VIP") + "имя категория VIP".length(),
//                            parsedText.indexOf("отчество номер полиса ДМС"));
//                    rawO = rawO.substring(0, rawO.indexOf( " "));
//                    fio += " " + rawO.replaceAll("\r|\n|[0-9]|[.!?\\-]|[ ]", "");
//
//                    System.out.println(fio);
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

        return "";
    }

}
