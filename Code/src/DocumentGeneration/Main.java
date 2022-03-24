package DocumentGeneration;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.text.DecimalFormat;

public class Main {

    public static void main(String[] args) {
        String[] details = {"J. Smith", "27 Sainsbury Close", "Stratford", "Essex EJ6 5TJ",
                "M", "DF65 POT", "2021-01-05"};
        String[] vehicle = {"DF65 POT", "Opel", "Vectra Estate"};
        String[][] jobDetails = {
                {"Exhaust pipe", "X784/6352J", "57.50", "1", "0"},
                {"Head Gasket", "Y76432-89T5", "15.75", "1", "0"},
                {"Valves", "672351X/34K", "5.15", "6", "0"},
                {"Exhaust pipe", "X784/6352J", "57.50", "3", "0"}};
        String[] mechanicDetails = {"105.00", "5.75"};
        String[] work = {"Replace exhaust", "Strip head and replace worn valves", "Replace grommet seals "};

        //MOTReminder moe = new MOTReminder();
        JobInvoice mot = new JobInvoice();
        //moe.generate(details);
        mot.generate(details,vehicle,work,jobDetails,mechanicDetails);
        System.out.println("Creation Successful");
    }
}