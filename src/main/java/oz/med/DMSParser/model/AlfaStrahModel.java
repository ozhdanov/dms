package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class AlfaStrahModel {

    String policyNumber;
    String fio;
    Date dateOfBirth;
    String adress;
    String phoneNumber;
    String placeOfWork;
    Date policyStartDate;
    Date policyEndDate;
    String policyType;
    Date deattachDate;
    boolean isNew = true;

}
