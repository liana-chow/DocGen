package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.text.DecimalFormat;
import java.time.LocalDate;

public class StockReport {
    public static void generate(String userName, String beginDate,
                                String endDate, String[][] stockDetails) {
        //Create a Document object
        Document document = new Document();
        DecimalFormat df = new DecimalFormat("0.00");

        String[] tableHeaders = {
                "Part Name", "Code", "Manufacturer", "Vehicle Type",
                "Year(s)", "Price, £", "Initial Stock Level", "Initial Cost, £",
                "Used", "Delivery", "New Stock level", "Stock Cost, £", "Low Level Threshold"};

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
        reportTitle.appendText( "Spare Parts / Stock Level Report" + "\n");

        Paragraph reportPeriod = section.addParagraph();
        reportPeriod.appendText( "Report Period: " + beginDate + " - " + endDate + "\n");

        double totalInitialCost= 0;
        double totalStockCost= 0;

        int tableLength = stockDetails.length + 5;
        //Add a table
        Table table = section.addTable(true);
        table.resetCells(tableLength-1, 13);

        TableRow row = table.getRows().get(0);
        row.setHeight(105f);
        for (int x = 0; x < tableHeaders.length; x++) {
            TableCell cell = row.getCells().get(x);
            cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            cell.getCellFormat().setTextDirection(TextDirection.Left_To_Right);
            cell.getCellFormat().getBorders().setLineWidth(2.5f);
            Paragraph p = cell.addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange txtRange = p.appendText(tableHeaders[x]);
            txtRange.getCharacterFormat().setFontSize(10f);
            txtRange.getCharacterFormat().setFontName("Arial");
            txtRange.getCharacterFormat().setBold(true);
        }

        for (int x = 0; x < stockDetails.length; x++){
            row = table.getRows().get(2+x);
            for (int y = 0; y < tableHeaders.length; y++) {
                TableCell cell = row.getCells().get(y);
                cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                Paragraph p = cell.addParagraph();
                p.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                TextRange txtRange;
                int newStock = Integer.parseInt(stockDetails[x][6])-Integer.parseInt(stockDetails[x][7])
                        +Integer.parseInt(stockDetails[x][8]);
                switch(y){
                    case 7:
                        double cost = Double.parseDouble(stockDetails[x][5])
                                *Double.parseDouble(stockDetails[x][6]);
                        totalInitialCost = totalInitialCost + cost;
                        txtRange = p.appendText(df.format(cost));
                        break;
                    case 10:
                        txtRange = p.appendText(String.valueOf(newStock));
                        break;
                    case 11:
                        double stockCost = newStock*Double.parseDouble(stockDetails[x][5]);
                        totalStockCost = totalStockCost + stockCost;
                        txtRange = p.appendText(df.format(stockCost));
                        break;
                    case 12:
                        txtRange = p.appendText(stockDetails[x][9]);
                        break;
                    default:
                        txtRange = p.appendText(stockDetails[x][y]);}
                txtRange.getCharacterFormat().setFontSize(10f);
                txtRange.getCharacterFormat().setFontName("Arial");
            }
        }

        table.getLastRow().getRowFormat().getBorders().setLineWidth(2.5f);
        Paragraph p =  table.get(tableLength-2,0).addParagraph();
        TextRange txtRange = p.appendText("Total");
        txtRange.getCharacterFormat().setFontSize(10f);
        txtRange.getCharacterFormat().setFontName("Arial");

        Paragraph t =  table.get(tableLength-2,7).addParagraph();
        TextRange txt = t.appendText(df.format(totalInitialCost));
        txt.getCharacterFormat().setFontSize(10f);
        txt.getCharacterFormat().setFontName("Arial");

        Paragraph l =  table.get(tableLength-2,11).addParagraph();
        TextRange txts = l.appendText(df.format(totalStockCost));
        txts.getCharacterFormat().setFontSize(10f);
        txts.getCharacterFormat().setFontName("Arial");

        table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Window);
        table.applyStyle(DefaultTableStyle.Table_Classic_1);

        Paragraph reportDate = section.addParagraph();
        reportDate.appendText("\n"+ "\n" + "Report Date: " +
                Formatting.getDayMonthYear(String.valueOf(LocalDate.now())));

        Paragraph genBy = section.addParagraph();
        genBy.appendText("\n"+ "\n" + "Generated by: " + "\n" + userName);


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
        genBy.applyStyle("paraStyle");

        //Save the document
        String fileName = java.time.LocalDate.now() + " Stock Report" + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}
