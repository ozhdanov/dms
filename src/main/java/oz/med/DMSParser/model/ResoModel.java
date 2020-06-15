package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class ResoModel {

    String policyNumber;
    String fio;
    String dateOfBirth;
    String adress;
    String phoneNumber;
    String placeOfWork;
    String validity;
    String policyType;
    boolean isNew = true;

}
