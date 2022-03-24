package DocumentGeneration;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.text.DecimalFormat;

public class Main {

    public static void main(String[] args) {
        DecimalFormat df = new DecimalFormat("0.00");

        String[] tableHeaders = {
                "Item                   ",
                "Part No.                   ",
                "Unit Cost  ",
                "Qty. ",
                "Cost (£)"};
        String[][] jobDetails = {
                {"Exhaust pipe", "X784/6352J", "57.50", "1", "0"},
                {"Head Gasket", "Y76432-89T5", "15.75", "1", "0"},
                {"Valves", "672351X/34K", "5.15", "6", "0"},
                {"Exhaust pipe", "X784/6352J", "57.50", "3", "0"}};
        String[] mechanicDetails = {"105.00", "5.75"};


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

        //Create a Document object
        Document document = new Document();

        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("paraStyle");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style1);

        //Add a section
        Section section = document.addSection();

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

        /*/dd data to the rest of rows
        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
            dataRow.setHeightType(TableRowHeightType.Exactly);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < data[r].length; c++) {
                dataRow.getCells().get(c).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                dataRow.getCells().get(c).addParagraph().appendText(data[r][c]);
            }
        }*/

        //Set background color for cells
        /*for (int j = 1; j < table.getRows().getCount(); j++) {
            if (j % 2 == 0) {
                TableRow row2 = table.getRows().get(j);
                for (int f = 0; f < row2.getCells().getCount(); f++) {
                    row2.getCells().get(f).getCellFormat().setBackColor(new Color(173, 216, 230));
                }
            }
        }*/

        //table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
        //Save to file
        document.saveToFile("CreateTable.docx", FileFormat.Docx_2013);
        System.out.println("Creation Successful");
    }
}