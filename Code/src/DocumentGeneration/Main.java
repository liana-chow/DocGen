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
        String[] comp = {"Fjord Distribution Ltd", "25 The Causeway, Staines, Middlesex"};
        String[][] stockreord = {
                {"Engine Block", "X93456", "Fjord", "Krapa", "2011-2015", "11.52", "8", "0", "3", "10"},
                {"Radiator", "C4563 ", "Fjord", "Ofcorsa", "2010-2021", "36.25", "7", "0", "0", "10"}};

        String[][] reports = {
                {"105", "11.30", "MOT", "Liana"},
                {"35", "12.30", "Annual Service", "Mo"},
                {"15", "13.30", "MOT", "Atahan"},
                {"125", "11.30", "MOT", "Atahan"},
                {"65", "16.30", "Repairs", "Atahan"},
                {"40", "20.30", "MOT", "Liana"},
                {"45", "30.30", "MOT", "Mo"},
                {"145", "40.30", "Annual Service", "Jeremy"},
                {"15", "50.30", "MOT", "Jeremy"},
                {"20", "10.30", "Repairs", "Vlad"},
                {"20", "10.30", "Repairs", "Ameen"},
                {"200", "13.30", "Other", "Vlad"},
                {"15", "15.30", "Annual Service", "Liana"},
        };

        String[][] monkreports = {
                {"Account Holder", "MOT"},
                {"Casual", "Repairs"},
                {"Casual", "Annual Service"},
                {"Account Holder", "Other"},
                {"Casual", "MOT"},
                {"Account Holder", "MOT"},
                {"Account Holder", "MOT"},
                {"Casual", "MOT"},
                {"Account Holder", "Repairs"},
                {"Account Holder", "MOT"},
                {"Account Holder", "Annual Service"},
        };



        //MOTReminder moe = new MOTReminder();
        //JobInvoice mot = new JobInvoice();
        //moe.generate(details);
        //mot.generate(details,vehicle,work,jobDetails,mechanicDetails);
        //StockReport stonk = new StockReport();
        //stonk.generate("Liana","01/11/2021", "30/11/2021", stock);
        //PartsOrder pant = new PartsOrder();
        //pant.generate("Liana",comp, stockreord);
        //JobReport jonk = new JobReport();
        //MonthlyReport monk = new MonthlyReport();
        //monk.generate("20-03-2022", "20-04-2022", monkreports);
        JobReport jronk = new JobReport();
        jronk.generate("20-03-2022", "20-04-2022", reports);
        System.out.println("Creation Successful");
    }
}