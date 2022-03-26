package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.text.DecimalFormat;
import java.time.LocalDate;

public class PartsOrder {
    public static void generate(String userName, String[] companyDetails, String[][] stockDetails) {
        //Create a Document object
        Document document = new Document();
        DecimalFormat df = new DecimalFormat("0.00");

        String[] tableHeaders = {"Order Number", "Description", "Quantity"};

        //Add a section
        Section section = document.addSection();

        //Add 3 paragraphs to the section
        Paragraph addressFormat = section.addParagraph();
        addressFormat.appendText("\n" + "\n" + "\n" + "\n" +
                "Quick Fix Fitters," + "\n" +
                "19 High St.," + "\n" +
                "Ashford," + "\n" +
                "Kent, CT16 8YY" + "\n");

        Paragraph writtenDate = section.addParagraph();
        writtenDate.appendText("Date: " + Formatting.getDayMonthYear(String.valueOf(LocalDate.now()))
                + "\n" + "\n");

        Paragraph companyAddressing = section.addParagraph();
        companyAddressing.appendText("Company: " + companyDetails[0] + "\n" +
                "Address: " + companyDetails[1] + "\n" + "\n" + "\n");

        Paragraph contactDetails = section.addParagraph();
        contactDetails.appendText("Tel: 01784 407862\n" +
                "Fax: 01784 407863" + "\n\n");

        double totalStockCost = 0;

        int tableLength = stockDetails.length + 4;
        //Add a table
        Table table = section.addTable(false);
        table.resetCells(tableLength, 4);

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

        //{0partName, 1code, 2manufacturer, 3vehicle type, 4years,
        // 5price, 6initial stock level, 7used, 8delivery, 9low level}}
        for (int x = 0; x < stockDetails.length; x++){
            row = table.getRows().get(2+x);
            int reorderAmount = Integer.parseInt(stockDetails[x][9])-
                    Integer.parseInt(stockDetails[x][6])+5;
            double reorderCost = Double.parseDouble(stockDetails[x][5])*Double.parseDouble(stockDetails[x][6]);
            totalStockCost = totalStockCost + reorderCost;
            for (int y = 0; y < tableHeaders.length+1; y++) {
                TableCell cell = row.getCells().get(y);
                cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph p = cell.addParagraph();
                p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                TextRange txtRange = null;
                switch(y) {
                    case 0:
                        txtRange = p.appendText(stockDetails[x][1]);
                        break;
                    case 1:
                        txtRange = p.appendText(stockDetails[x][0]);
                        break;
                    case 2:
                        txtRange = p.appendText(String.valueOf(reorderAmount));
                        break;
                    case 3:
                        txtRange = p.appendText("£" + df.format(reorderCost));
                        break;
                }
                txtRange.getCharacterFormat().setFontSize(11f);
                txtRange.getCharacterFormat().setFontName("Arial");
            }
        }

        Paragraph p = table.get(tableLength-1,2).addParagraph();
        TextRange txtRange = p.appendText("Total: ");
        txtRange.getCharacterFormat().setFontSize(11f);
        txtRange.getCharacterFormat().setFontName("Arial");

        Paragraph l = table.get(tableLength-1,3).addParagraph();
        TextRange txtR = l.appendText("£" + df.format(totalStockCost));
        txtR.getCharacterFormat().setFontSize(11f);
        txtR.getCharacterFormat().setFontName("Arial");

        table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Window);

        Paragraph genBy = section.addParagraph();
        genBy.appendText("\n" + "\n" + "Signed: " + "\n" + userName);


        //Set title style for paragraph 1
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("paraStyle");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style1);

        addressFormat.applyStyle("paraStyle");
        writtenDate.applyStyle("paraStyle");
        companyAddressing.applyStyle("paraStyle");
        contactDetails.applyStyle("paraStyle");
        genBy.applyStyle("paraStyle");

        //Save the document
        String fileName = java.time.LocalDate.now() + " " + companyDetails[0] + " Stock Order" + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}
