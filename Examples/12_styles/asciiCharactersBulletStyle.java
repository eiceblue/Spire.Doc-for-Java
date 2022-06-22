import com.spire.doc.*;
import com.spire.doc.documents.*;

/**
 * Create bullet styles by ASCII characters
 */
public class asciiCharactersBulletStyle {
    public static void main(String[] args) {
        //Create a new document
        Document document = new Document();
        Section section = document.addSection();

        //Create four list styles based on different ASCII characters
        ListStyle listStyle1 = new ListStyle(document, ListType.Bulleted);
        listStyle1.setName("listStyle");
        listStyle1.getLevels().get(0).setBulletCharacter(ASCII2String(0x006e));
        listStyle1.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
        document.getListStyles().add(listStyle1);

        ListStyle listStyle2 = new ListStyle(document, ListType.Bulleted);
        listStyle2.setName("listStyle2");
        listStyle2.getLevels().get(0).setBulletCharacter(ASCII2String(0x0075));
        listStyle2.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
        document.getListStyles().add(listStyle2);

        ListStyle listStyle3 = new ListStyle(document, ListType.Bulleted);
        listStyle3.setName("listStyle3");
        listStyle3.getLevels().get(0).setBulletCharacter(ASCII2String(0x00b2));
        listStyle3.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
        document.getListStyles().add(listStyle3);

        ListStyle listStyle4 = new ListStyle(document, ListType.Bulleted);
        listStyle4.setName("listStyle4");
        listStyle4.getLevels().get(0).setBulletCharacter(ASCII2String(0x00d8));
        listStyle4.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
        document.getListStyles().add(listStyle4);

        //Add four paragraphs and apply list style separately
        Paragraph p1 = section.getBody().addParagraph();
        p1.appendText("Spire.Doc for Java");
        p1.getListFormat().applyStyle(listStyle1.getName());
        Paragraph p2 = section.getBody().addParagraph();
        p2.appendText("Spire.Doc for Java");
        p2.getListFormat().applyStyle(listStyle2.getName());
        Paragraph p3 = section.getBody().addParagraph();
        p3.appendText("Spire.Doc for Java");
        p3.getListFormat().applyStyle(listStyle3.getName());
        Paragraph p4 = section.getBody().addParagraph();
        p4.appendText("Spire.Doc for Java");
        p4.getListFormat().applyStyle(listStyle4.getName());

        //Save the document
        String output = "output/ASCIICharactersBulletStyle_output.docx";
        document.saveToFile(output, FileFormat.Docx);
    }

    public static String ASCII2String(int ascii) {
        return String.valueOf((char) ascii);
    }
}
