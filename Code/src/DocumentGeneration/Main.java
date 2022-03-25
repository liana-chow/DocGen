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

        String[][] stock = {
                {"Grommet", "X66745877", "Fjord", "Krapa", "2011-2015", "0.90", "34", "2", "0", "10"},
                {"", "D43-78 ", "Vauxhill", "Ofcorsa", "2010-2021", "1.30", "12", "0", "0", "10"},
                {"Water Pump", "G457", "Fjord", "Krapa", "2010-2014", "56.70", "8", "0", "3", "10"},
                {"", "H456-9UI", "Volva", "S34", "2008-2015", "124.34", "2", "0", "0", "3"}};
        //{0partName, 1code, 2manufacturer, 3vehicle type, 4years,
        // 5price, 6initial stock level, 7used, 8delivery, 9low level}}



        //MOTReminder moe = new MOTReminder();
        //JobInvoice mot = new JobInvoice();
        //moe.generate(details);
        //mot.generate(details,vehicle,work,jobDetails,mechanicDetails);
        StockReport stonk = new StockReport();
        stonk.generate("Liana","01/11/2021", "30/11/2021", stock);
        System.out.println("Creation Successful");
    }
}