package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;
import java.text.DecimalFormat;
import java.time.LocalDate;

public class JobInvoice {
    public static void generate(String[] customerDetails, String[] vehicleDetails,
                                String[] description, String[][] jobDetails, String[] mechanicDetails) {
        //Create a Document object
        Document document = new Document();
        DecimalFormat df = new DecimalFormat("0.00");

        String[] tableHeaders = {
                "Item                   ",
                "Part No.                   ",
                "Unit Cost  ",
                "Qty. ",
                "Cost (£)"};

        //Add a section
        Section section = document.addSection();

        //Add 3 paragraphs to the section
        Paragraph addressFormat = section.addParagraph();
        addressFormat.appendText("\n" + "\n" + "\n" + "\n" +
                customerDetails[0] + ",\n" +
                customerDetails[1] + ",\n" +
                customerDetails[2] + ",\n" +
                customerDetails[3] +",\n");

        Formatting.addressFormatting(document);

        Paragraph writtenDate = section.addParagraph();
        writtenDate.appendText(Formatting.getDayMonthYear(String.valueOf(LocalDate.now())));


        Paragraph toCustomer = section.addParagraph();
        toCustomer.appendText("Dear " + Formatting.addressingCustomer(customerDetails[4]) + customerDetails[0].substring(3));

        Paragraph letterTitle = section.addParagraph();
        letterTitle.appendText("\n" + "INVOICE NO.: " + customerDetails[0] + "\n" + "\n");

        Paragraph vehicleBody = section.addParagraph();
        vehicleBody.appendText("Vehicle Registration No.: " + vehicleDetails[0] + "\n" +
                "Make: " + vehicleDetails[1] + "\n" +
                "Model: " + vehicleDetails[2] + "\n");

        Paragraph descriptionWork = section.addParagraph();
        String formattedDescription = Formatting.formatDescription(description);
        descriptionWork.appendText("Description of work: " + "\n" + formattedDescription + "\n");

        double totalUnitCostInt= 0;
        for (int x = 0; x < jobDetails.length; x++) {
            double cost = Double.parseDouble(jobDetails[x][2]) * Double.parseDouble(jobDetails[x][3]);
            totalUnitCostInt = totalUnitCostInt + cost;
            jobDetails[x][4] = df.format(cost);
        }
        String totalLabour= df.format(Double.parseDouble(mechanicDetails[0])*Double.parseDouble(mechanicDetails[1]));
        String totalJob= df.format(totalUnitCostInt+Double.parseDouble(totalLabour));
        String totalVAT= df.format(Double.parseDouble(totalJob)*1.2);
        String grandTotal= df.format(Double.parseDouble(totalJob)+Double.parseDouble(totalVAT));

        int tableLength = jobDetails.length + 13;
        //Add a table
        Table table = section.addTable(false);
        table.resetCells(tableLength-1, 5);

        TableRow row = table.getRows().get(0);
        for (int x = 0; x < tableHeaders.length; x++) {
            TableCell cell = row.getCells().get(x);
            cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            cell.getCellFormat().getBorders().getBottom().setBorderType(BorderStyle.Single);
            Paragraph p = cell.addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
            TextRange txtRange = p.appendText(tableHeaders[x]);
            txtRange.getCharacterFormat().setFontSize(11f);
            txtRange.getCharacterFormat().setFontName("Arial");
        }

        for (int x = 0; x < jobDetails.length; x++){
            row = table.getRows().get(2+x);
            for (int y = 0; y < jobDetails[0].length; y++) {
                TableCell cell = row.getCells().get(y);
                cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph p = cell.addParagraph();
                p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                TextRange txtRange = p.appendText(jobDetails[x][y]);
                txtRange.getCharacterFormat().setFontSize(11f);
                txtRange.getCharacterFormat().setFontName("Arial");
            }
        }

        String[][] informationFormatting = {
                {"10","0", "Labour"},
                {"10","2", mechanicDetails[0]},
                {"10","3", mechanicDetails[1]},
                {"10","4", totalLabour},
                {"7","2", "Total"},
                {"7","4", totalJob},
                {"5","2", "VAT"},
                {"5","4", totalVAT},
                {"2","2", "Grand Total"},
                {"2","4", String.valueOf(grandTotal)},

        };

        for (String[] strings : informationFormatting) {
            int x = tableLength - Integer.parseInt(strings[0]);
            int y = Integer.parseInt(strings[1]);
            TableCell cell = table.get(x,y);
            cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = cell.addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
            TextRange txtRange = p.appendText(strings[2]);
            txtRange.getCharacterFormat().setFontSize(11f);
            txtRange.getCharacterFormat().setFontName("Arial");
        }

        table.get(tableLength-4,3).getCellFormat().getBorders().
                getBottom().setBorderType(BorderStyle.Single);

        table.get(tableLength-9,3).getCellFormat().getBorders().
                getBottom().setBorderType(BorderStyle.Single);

        table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Window);

        Paragraph thankCustomer = section.addParagraph();
        thankCustomer.appendText("\n"+ "\n" + "Thank you for your valued custom. " +
                "We look forward to receiving your payment in due course.  " + "\n");

        Formatting.signOff(document);

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
        writtenDate.applyStyle("paraStyle");
        toCustomer.applyStyle("paraStyle");
        letterTitle.applyStyle("boldStyle");
        vehicleBody.applyStyle("paraStyle");
        descriptionWork.applyStyle("paraStyle");
        thankCustomer.applyStyle("paraStyle");

        //Horizontally align paragraph 1 to center
        letterTitle.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        writtenDate.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

        //Save the document
        String fileName = java.time.LocalDate.now() + " Invoice Job " + customerDetails[0] + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}