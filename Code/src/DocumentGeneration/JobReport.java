package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.awt.*;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class JobReport {
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
        reportTitle.appendText("\n" + "Monthly Job Report" + "\n");

        Paragraph reportPeriod = section.addParagraph();
        reportPeriod.appendText( "Report Period: " + beginDate + " - " + endDate + "\n");

        Paragraph reportDate = section.addParagraph();
        reportDate.appendText( "Report Date: " +
                Formatting.getDayMonthYear(String.valueOf(LocalDate.now())) + "\n" + "\n");

        double totalTime = 0;
        double totalPrice = 0;
        ArrayList<String[]> MOTJobs = new ArrayList<>();
        ArrayList<String[]> annualJobs = new ArrayList<>();
        ArrayList<String[]> repairsJob = new ArrayList<>();
        ArrayList<String[]> otherJobs = new ArrayList<>();
        Map<String, ArrayList<String[]>> mechanicJob = new HashMap<>();
        for (String[] x : jobDetails){
                mechanicJob.put(x[3], new ArrayList<>());
        }

        for (String[] x : jobDetails){
            totalPrice = totalPrice + Double.parseDouble(x[1]);
            totalTime = totalTime + Double.parseDouble(x[0]);
            ArrayList<String[]> temp = mechanicJob.get(x[3]);
            temp.add(x);
            mechanicJob.put(String.valueOf(x[3]), temp);
            switch (x[2]) {
                case "MOT" -> {
                    MOTJobs.add(x);
                }
                case "Annual Service" -> {
                    annualJobs.add(x);
                }
                case "Repairs" -> {
                    repairsJob.add(x);
                }
                case "Other" -> {
                    otherJobs.add(x);
                }
            }
        }

        double MOTtimeTaken = 0;
        double MOTpriceOf = 0;
        for (String[] x : MOTJobs){
            MOTtimeTaken = MOTtimeTaken + Double.parseDouble(x[0]);
            MOTpriceOf = MOTpriceOf + Double.parseDouble(x[1]);
        }

        double AnnualtimeTaken = 0;
        double AnnualpriceOf = 0;
        for (String[] x : annualJobs){
            AnnualtimeTaken = AnnualtimeTaken + Double.parseDouble(x[0]);
            AnnualpriceOf = AnnualpriceOf + Double.parseDouble(x[1]);
        }

        double repairstimeTaken = 0;
        double repairspriceOf = 0;
        for (String[] x : repairsJob){
            repairstimeTaken = repairstimeTaken + Double.parseDouble(x[0]);
            repairspriceOf = repairspriceOf + Double.parseDouble(x[1]);
        }

        double OthertimeTaken = 0;
        double OtherpriceOf = 0;
        for (String[] x : otherJobs){
            OthertimeTaken = OthertimeTaken + Double.parseDouble(x[0]);
            OtherpriceOf = OtherpriceOf + Double.parseDouble(x[1]);
        }
        String[] overallAverages = {
                df.format(totalTime/jobDetails.length), df.format(totalPrice/jobDetails.length),
                df.format(MOTtimeTaken/MOTJobs.size()), df.format(MOTpriceOf/MOTJobs.size()),
                df.format(AnnualtimeTaken/annualJobs.size()), df.format(AnnualpriceOf/annualJobs.size()),
                df.format(repairstimeTaken/repairsJob.size()), df.format(repairspriceOf/repairsJob.size()),
                df.format(OthertimeTaken/otherJobs.size()), df.format(OtherpriceOf/otherJobs.size())
        };
        for (int i = 1; i < overallAverages.length; i++){
            if (totalTime==0){
                overallAverages[0] = "N/A";
                overallAverages[1] = "N/A";
            }
            if (MOTtimeTaken==0){
                overallAverages[2] = "N/A";
                overallAverages[3] = "N/A";
            }
            if (AnnualtimeTaken==0){
                overallAverages[4] = "N/A";
                overallAverages[5] = "N/A";
            }
            if (repairstimeTaken==0){
                overallAverages[6] = "N/A";
                overallAverages[7] = "N/A";
            }
            if (OthertimeTaken==0){
                overallAverages[8] = "N/A";
                overallAverages[9] = "N/A";
            }
        }
        Paragraph totalAverage = section.addParagraph();
        totalAverage.appendText("Total Average Time: " + overallAverages[0] + "\n" +
                "Total Average Price: " + overallAverages[1] + "\n");

        Paragraph MOTAverages = section.addParagraph();
        MOTAverages.appendText("MOT Average Time: " + overallAverages[2] + "\n" +
                "MOT Average Price: " + overallAverages[3] + "\n");

        Paragraph annualAverages = section.addParagraph();
        annualAverages.appendText("Annual Service Average Time: " + overallAverages[4] + "\n" +
                "Annual Service Average Price: " + overallAverages[5] + "\n");

        Paragraph repairAverages = section.addParagraph();
        repairAverages.appendText("Repairs Average Time: " + overallAverages[6] + "\n" +
                "Repairs Average Price: " + overallAverages[7] + "\n");

        Paragraph otherAverages = section.addParagraph();
        otherAverages.appendText("Other Jobs Average Time: " + overallAverages[8]
                + "\n" + "Other Jobs Average Price: " + overallAverages[9] + "\n");

        ArrayList<String[]> mechanicAverages = new ArrayList<>();
        for (var entry : mechanicJob.entrySet()) {
            double timeTaken = 0; double priceOf = 0; int counter = 0;
            double motTime = 0; double motPrice = 0; int doneMOT = 0;
            double annualTime = 0; double annualPrice = 0; int doneAnnual = 0;
            double repairsTime = 0; double repairsPrice = 0; int doneRepair = 0;
            double otherTime = 0; double otherPrice = 0; int doneOther = 0;

            String key = entry.getKey();
            ArrayList<String[]> value = entry.getValue();
            for (String[] strings : value) {
                counter++;
                timeTaken = timeTaken + Double.parseDouble(strings[0]);
                priceOf = priceOf + Double.parseDouble(strings[1]);
                switch (strings[2]) {
                    case "MOT" -> {
                        motTime = motTime + Double.parseDouble(strings[0]);
                        motPrice = motPrice + Double.parseDouble(strings[1]);
                        doneMOT++;
                    }
                    case "Annual Service" -> {
                        annualTime = annualTime + Double.parseDouble(strings[0]);
                        annualPrice = annualPrice + Double.parseDouble(strings[1]);
                        doneAnnual++;
                    }
                    case "Repairs" -> {
                        repairsTime = repairsTime + Double.parseDouble(strings[0]);
                        repairsPrice = repairsPrice + Double.parseDouble(strings[1]);
                        doneRepair++;
                    }
                    case "Other" -> {
                        otherTime = otherTime + Double.parseDouble(strings[0]);
                        otherPrice = otherPrice + Double.parseDouble(strings[1]);
                        doneOther++;
                    }
                }
            }
            String[] averages = {key, String.valueOf(counter),
                    df.format(timeTaken / counter), df.format(priceOf / counter),
                    df.format(motTime / doneMOT), df.format(motPrice / doneMOT),
                    df.format(annualTime / doneAnnual), df.format(annualPrice / doneAnnual),
                    df.format(repairsTime / doneRepair), df.format(repairsPrice / doneRepair),
                    df.format(otherTime / doneOther), df.format(otherPrice / doneOther)};
            for (int i = 1; i < averages.length; i++){
                if (counter==0){
                    averages[2] = "N/A";
                    averages[3] = "N/A";
                }
                if (doneMOT==0){
                    averages[4] = "N/A";
                    averages[5] = "N/A";
                }
                if (doneAnnual==0){
                    averages[6] = "N/A";
                    averages[7] = "N/A";
                }
                if (doneRepair==0){
                    averages[8] = "N/A";
                    averages[9] = "N/A";
                }
                if (doneOther==0){
                    averages[10] = "N/A";
                    averages[11] = "N/A";
                }
            }
            mechanicAverages.add(averages);
        }

        String[] tableHeaders = {"Mechanic", "Taken Jobs",
                "Overall Time Avg", "Overall Price Avg",
                "MOT Jobs Time Avg", "MOT Jobs Price Avg",
                "Annual Service Time Avg", "Annual Service Price Avg",
                "Repair Jobs Time Avg", "Repair Jobs Price Avg",
                "Other Jobs Time Avg", "Other Jobs Price Avg"};

        int tableLength = mechanicAverages.size() + 1;
        //Add a table
        Table table = section.addTable(true);
        table.resetCells(tableLength, 12);

        TableRow row = table.getRows().get(0);
        row.setHeight(120f);
        for (int x = 0; x < tableHeaders.length; x++) {
            TableCell cell = row.getCells().get(x);
            cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            cell.getCellFormat().setTextDirection(TextDirection.Left_To_Right);
            cell.getCellFormat().getBorders().setLineWidth(1.5f);
            Paragraph p = cell.addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange txtRange = p.appendText(tableHeaders[x]);
            txtRange.getCharacterFormat().setFontSize(10f);
            txtRange.getCharacterFormat().setFontName("Arial");
            txtRange.getCharacterFormat().setBold(true);
        }

        Color notaval = new Color(255, 175,175);
        for (int x = 0; x < mechanicAverages.size(); x++){
            row = table.getRows().get(1+x);
            String[] temp = mechanicAverages.get(x);
            for (int y = 0; y < tableHeaders.length; y++) {
                TableCell cell = row.getCells().get(y);
                cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph p = cell.addParagraph();
                p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                TextRange txtRange = p.appendText(temp[y]);
                txtRange.getCharacterFormat().setFontSize(11f);
                txtRange.getCharacterFormat().setFontName("Arial");
                if (temp[y]=="N/A"){
                    cell.getCellFormat().setBackColor(notaval);
                }
            }
        }

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
        totalAverage.applyStyle("paraStyle");
        MOTAverages.applyStyle("paraStyle");
        annualAverages.applyStyle("paraStyle");
        repairAverages.applyStyle("paraStyle");
        otherAverages.applyStyle("paraStyle");


        //Save the document
        String fileName = java.time.LocalDate.now() + " Job Report" + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}
