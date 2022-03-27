package DocumentGeneration;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;

import java.text.DecimalFormat;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

public class MonthlyReport {
    public static void generate(String beginDate, String endDate, String[][] jobDetails) {
        //Create a Document object
        Document document = new Document();
        DecimalFormat df = new DecimalFormat("0.00");

        //Add a section
        Section section = document.addSection();

        //Add 3 paragraphs to the section
        Paragraph addressFormat = section.addParagraph();
        addressFormat.appendText("\n" + "\n" + "\n" + "\n" +
                "Quick Fix Fitters," + "\n" +
                "19 High St.," + "\n" +
                "Ashford," + "\n" +
                "Kent, CT16 8YY" + "\n");

        Paragraph reportTitle = section.addParagraph();
        reportTitle.appendText( "\n" + "\n" + "Monthly Job Report");

        Paragraph reportPeriod = section.addParagraph();
        reportPeriod.appendText( "Report Period: " + beginDate + " - " + endDate + "\n");

        Paragraph reportDate = section.addParagraph();
        reportDate.appendText("\n"+ "Report Date: " +
                Formatting.getDayMonthYear(String.valueOf(LocalDate.now())));

        Map<String, Integer> bookedJobs = new HashMap<String, Integer>();
        bookedJobs.put("MOT", 0);
        bookedJobs.put("Annual Service", 0);
        bookedJobs.put("Repairs", 0);
        bookedJobs.put("Other", 0);
        int accountHolderCustomer = 0;
        int casualCustomer = 0;
        int totalJobs = 0;
        for (String[] x : jobDetails){
            if (x[0] == "Account Holder"){
                accountHolderCustomer = accountHolderCustomer + 1;
            } else {
                casualCustomer = casualCustomer + 1;
            }
            totalJobs = totalJobs + 1;
            switch (x[1]){
                case "MOT":
                    bookedJobs.put("MOT", bookedJobs.get("MOT") + 1);
                    break;
                case "Annual Service":
                    bookedJobs.put("Annual Service", bookedJobs.get("Annual Service") + 1);
                    break;
                case "Repairs":
                    bookedJobs.put("Repairs", bookedJobs.get("Repairs") + 1);
                    break;
                case "Other":
                    bookedJobs.put("Other", bookedJobs.get("Other") + 1);
                    break;
            }
        }

        Paragraph jobTypeHeader = section.addParagraph();
        jobTypeHeader.appendText("\n"+ "Jobs Booked: ");
        Paragraph numberOfVehicles = section.addParagraph();
        numberOfVehicles.appendText(
                "Total Jobs: " + totalJobs + "\n" +
                "MOT Jobs booked: " + bookedJobs.get("MOT") + "\n" +
                "Annual Service Jobs booked: " + bookedJobs.get("Annual Service") + "\n" +
                "Repairs Jobs booked: " + bookedJobs.get("Repairs") + "\n" +
                "Other Jobs booked: " + bookedJobs.get("Other") + "\n");

        Paragraph customerTypeHeader = section.addParagraph();
        customerTypeHeader.appendText("\n"+ "Customer Types: ");
        Paragraph customerType = section.addParagraph();
        customerType.appendText(
                "Account Holder Customers: " + accountHolderCustomer + "\n" +
                "Casual Customers: " + casualCustomer + "\n");


        //Set title style for paragraph 1
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("paraStyle");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style1);

        //Set style for paragraph 2 and 3
        ParagraphStyle style2 = new ParagraphStyle(document);
        style2.setName("boldStyle");
        style2.getCharacterFormat().setBold(true);
        style2.getCharacterFormat().setFontName("Arial");
        style2.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style2);

        addressFormat.applyStyle("paraStyle");
        reportTitle.applyStyle("boldStyle");
        reportPeriod.applyStyle("paraStyle");
        reportDate.applyStyle("paraStyle");
        numberOfVehicles.applyStyle("paraStyle");
        customerType.applyStyle("paraStyle");
        customerTypeHeader.applyStyle("boldStyle");
        jobTypeHeader.applyStyle("boldStyle");

        //Save the document
        String fileName = java.time.LocalDate.now() + " Monthly Report" + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}
