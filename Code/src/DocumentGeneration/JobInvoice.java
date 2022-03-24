package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;

import java.lang.reflect.Array;
import java.time.LocalDate;

public class JobInvoice {
    public static void generate(String[] customerDetails, String[] vehicleDetails, String[] jobDetails) {
        //Create a Document object
        Document document = new Document();
        Formatting format = new Formatting();

        //Add a section
        Section section = document.addSection();

        //Add 3 paragraphs to the section
        Paragraph addressFormat = section.addParagraph();
        addressFormat.appendText("\n" + "\n" + "\n" + "\n" +
                (String) Array.get(customerDetails, 0) + ",\n" +
                (String) Array.get(customerDetails, 1) + ",\n" +
                (String) Array.get(customerDetails, 2) + ",\n" +
                (String) Array.get(customerDetails, 3) +",\n");

        format.addressFormatting(document);

        Paragraph writtenDate = section.addParagraph();
        writtenDate.appendText(format.getDayMonthYear(String.valueOf(LocalDate.now())));


        Paragraph toCustomer = section.addParagraph();
        toCustomer.appendText("Dear " + format.addressingCustomer(customerDetails[4]) + customerDetails[0].substring(3));

        Paragraph letterTitle = section.addParagraph();
        letterTitle.appendText("\n" + "INVOICE NO.: " + (String) Array.get(customerDetails, 0) + "\n" + "\n");

        Paragraph vehicleBody = section.addParagraph();
        vehicleBody.appendText("Vehicle Registration No.: " + (String) Array.get(vehicleDetails, 0) + "\n" +
                "Make: " + (String) Array.get(vehicleDetails, 1) + "\n" +
                "Model: " + (String) Array.get(vehicleDetails, 2) + "\n");

        format.signOff(document);

        //Set title style for paragraph 1
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("paraStyle");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(10f);
        document.getStyles().add(style1);

        //Set style for paragraph 2 and 3
        ParagraphStyle style2 = new ParagraphStyle(document);
        style2.setName("boldStyle");
        style2.getCharacterFormat().setBold(true);
        style2.getCharacterFormat().setFontName("Arial");
        style2.getCharacterFormat().setFontSize(10f);
        document.getStyles().add(style2);

        addressFormat.applyStyle("paraStyle");
        writtenDate.applyStyle("paraStyle");
        toCustomer.applyStyle("paraStyle");
        letterTitle.applyStyle("boldStyle");
        vehicleBody.applyStyle("paraStyle");

        //Horizontally align paragraph 1 to center
        letterTitle.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        writtenDate.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

        //Save the document
        String fileName = java.time.LocalDate.now() + " Invoice Job " + (String) Array.get(jobDetails, 0) + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}