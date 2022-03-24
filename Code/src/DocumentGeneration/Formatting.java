package DocumentGeneration;

import com.spire.doc.Document;
import com.spire.doc.Section;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextBox;
import com.spire.doc.fields.TextRange;

import java.awt.*;
import java.lang.reflect.Array;
import java.time.LocalDate;
import java.time.Month;

public class Formatting {
    static String addressingCustomer(String g){
        String addressingCustomer;
        if (g == "F") {
            addressingCustomer = "Ms. ";
        } else {
            addressingCustomer = "Mr. ";
        }
        return addressingCustomer;
    }

    static String formatDescription(String[] desc){
        String s = "";
        for (int i = 0; i < desc.length; i++){
            s = s + (i + 1) + ") " + desc[i] + "\n";
        }
        return s;
    }

    static void signOff(Document document){
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("signStyle");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style1);

        Section section = document.getSections().get(0);
        Paragraph letterSignOff = section.addParagraph();
        letterSignOff.appendText("\n" + "\n" + "\n" + "\n" + "Yours sincerely," + "\n" + "\n" + "G. Lancaster");
        letterSignOff.applyStyle("signStyle");
    }

    static void addressFormatting(Document document){
        ParagraphStyle style1 = new ParagraphStyle(document);
        style1.setName("Style");
        style1.getCharacterFormat().setFontName("Arial");
        style1.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(style1);

        Section section = document.getSections().get(0);

        Paragraph addressQFF = section.addParagraph();
        TextBox tb = addressQFF.appendTextBox(100f, 60f);
        tb.getFormat().setHorizontalOrigin(HorizontalOrigin.Right_Margin_Area);
        tb.getFormat().setVerticalOrigin(VerticalOrigin.Page);

        //set the position of textbox
        tb.getFormat().setHorizontalOrigin(HorizontalOrigin.Right_Margin_Area);
        tb.getFormat().setHorizontalPosition(-100f);
        tb.getFormat().setVerticalOrigin(VerticalOrigin.Page);
        tb.getFormat().setVerticalPosition(60f);

        tb.getFormat().setLineColor(new Color(255,255,255));

        addressQFF = tb.getBody().addParagraph();
        TextRange textRange = addressQFF.appendText("Quick Fix Fitters," + "\n" +
                "19 High St.," + "\n" +
                "Ashford," + "\n" +
                "Kent, CT16 8YY" + "\n");
        textRange.getCharacterFormat().setFontName("Arial");
        textRange.getCharacterFormat().setFontSize(11f);
        addressQFF.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
    }

    static String getDayMonthYear(String date){
        // Get an instance of LocalTime from date
        LocalDate currentDate = LocalDate.parse(date);

        String ending;

        // Get day from date
        int day = currentDate.getDayOfMonth();
        // Get month from date
        Month month = currentDate.getMonth();
        // Get year from date
        int year = currentDate.getYear();

        switch (day) {
            case 1:
                ending = "st";
                break;
            case 2:
                ending = "nd";
                break;
            case 3:
                ending = "rd";
                break;
            default:
                ending = "th";
        }
        String higherMonth = month.name().substring(0,1);
        String lowerMonth = month.name().substring(1);
        // Print the day, month, and year
        String longDate = day + ending + " " + higherMonth + lowerMonth.toLowerCase() + " " + year;
        return longDate;
    }

    static String spacing(int length, int line){
        int spacing = 0;
        String space = "";
        switch (line) {
            case 0:
                spacing = 121 - length;
                break;
            case 1:
                spacing = 120 - length;
                break;
            case 2:
                spacing = 126 - length;
                break;
            case 3:
                spacing = 116 - length;
        }
        for (int i = 0; i < spacing; i++){
            space = space + " ";
        }
        return space;
    }
}
