package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class AbsolutModel {

    String policyNumber;
    String fio;
    String dateOfBirth;
    String adress;
    String validity;
    String insuranceProgram;
    String insurant;
    boolean isNew = true;

}
