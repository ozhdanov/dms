package oz.med.DMSParser.model;

import lombok.Data;

import java.util.Date;

@Data
public class BestDoctorModel {

    String policyNumber;
    String surname;
    String name;
    String patronymic;
    String sex;
    Date dateOfBirth;
    String adress;
    String phoneNumber;
    String placeOfWork;
    Date policyStartDate;
    Date policyEndDate;
    boolean isNew = true;

}
