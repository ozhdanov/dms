package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class RosGosStrahModel {

    String policyNumber;
    String fio;
    Date dateOfBirth;
    String adress;
    String phoneNumber;
    String sex;
    Date deattachDate;
    boolean isNew = true;

}
