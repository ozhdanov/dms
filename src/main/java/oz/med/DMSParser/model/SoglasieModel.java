package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class SoglasieModel {

    String policyNumber;
    String surname;
    String name;
    String patronymic;
    String dateOfBirth;
    String adress;
    String phoneNumber;
    String placeOfWork;
    String validity;
    boolean isNew = true;

}
