package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class InGosStrahModel {

    String policyNumber;
    String surname;
    String name;
    String patronymic;
    Date dateOfBirth;
    String sex;
    String adressAndPhoneNumber;
    String note;
    String plan;
    String insuranceProgram;
    Date policyStartDate;
    Date policyEndDate;
    String insuranceNote;
    String insuranceExtension;
    String limitations;
    String characteristic;
    boolean isNew = true;

}
