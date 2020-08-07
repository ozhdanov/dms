package oz.med.DMSParser.model;

import lombok.Data;

@Data
public class RenaissanceModel {

    String fio;
    String dateOfBirth;
    String passport;
    String adress;
    String phoneNumber;
    String policyNumber;
    String placeOfWork;
    String validity;
    boolean isNew = true;

}
