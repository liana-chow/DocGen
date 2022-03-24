package DocumentGeneration;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextBox;
import com.spire.doc.fields.TextRange;

import java.awt.*;
import java.lang.reflect.Array;
import java.text.Format;
import java.time.LocalDate;

public class MOTReminder {
    public void generate(String[] details) {
        //Create a Document object
        Document document = new Document();
        Formatting format = new Formatting();

        //Add a section
        Section section = document.addSection();

        //Add 3 paragraphs to the section
        Paragraph addressFormat = section.addParagraph();
        addressFormat.appendText("\n" + "\n" + "\n" + "\n" +
                (String) Array.get(details, 0) + ",\n" +
                (String) Array.get(details, 1) + ",\n" +
                (String) Array.get(details, 2) + ",\n" +
                (String) Array.get(details, 3) +",\n");

        format.addressFormatting(document);

        Paragraph writtenDate = section.addParagraph();
        writtenDate.appendText(format.getDayMonthYear(String.valueOf(LocalDate.now())));


        Paragraph toCustomer = section.addParagraph();
        toCustomer.appendText("Dear " + format.addressingCustomer(details[4]) + details[0].substring(3));

        Paragraph letterTitle = section.addParagraph();
        letterTitle.appendText("\n" + "REMINDER - MoT TEST DUE" + "\n" + "\n" +
                "Vehicle Registration No.: " + (String) Array.get(details, 5) + "     " +
                "Renewal Test: " + format.getDayMonthYear((String) Array.get(details, 6)) + "\n");

        Paragraph letterBody = section.addParagraph();
        letterBody.appendText("According to our records, the above vehicle is due to have its MoT certificate renewed on the date shown.\n" +
                "\n" + "Account Holders customers such as yourself are assured of our prompt attention, and we hope that you will use our " +
                "services on this occasion in order to have the necessary test carried out on your vehicle.");
        letterBody.getFormat().setWordWrap(true);

        format.signOff(document);

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
        letterBody.applyStyle("paraStyle");

        //Horizontally align paragraph 1 to center
        letterTitle.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        writtenDate.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

        //Save the document
        String fileName = java.time.LocalDate.now() + " MOT Reminder " + (String) Array.get(details, 0) + ".docx";
        document.saveToFile(fileName, FileFormat.Docx);
    }
}