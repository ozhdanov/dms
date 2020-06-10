package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class SogazModel {

    String policyNumber;
    String surname;
    String name;
    String patronymic;
    Date dateOfBirth;
    String sex;
    String adress;
    String homePhoneNumber;
    String workPhoneNumber;
    String mobilPhoneNumber;
    String insuranceProgram;
    Date policyStartDate;
    Date policyEndDate;
    String insuranceNote;
    String insuranceExtension;
    String placeOfWork;
    String position;
    boolean isNew = true;

}
