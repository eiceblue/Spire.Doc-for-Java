import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.Field;

import java.util.ArrayList;

public class modifyHyperlinkText {
    public static void main(String[] args) {
        //Load Document
        String input = "data/JAVAHyperlinksTemp_N.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Find all hyperlinks in the Word document
        ArrayList<Field> hyperlinks = new ArrayList<Field>();
        for (Section section : (Iterable<Section>)doc.getSections())
        {
            for (DocumentObject object :  (Iterable<DocumentObject>)section.getBody().getChildObjects())
            {
                if (object.getDocumentObjectType().equals(DocumentObjectType.Paragraph))
                {
                    Paragraph paragraph=(Paragraph)object;
                    for (DocumentObject cObject : (Iterable<DocumentObject>)paragraph.getChildObjects())
                    {
                        if (cObject.getDocumentObjectType().equals(DocumentObjectType.Field))
                        {
                            Field field = (Field)cObject;
                            if (field.getType().equals( FieldType.Field_Hyperlink))
                            {
                                hyperlinks.add(field);
                            }
                        }
                    }
                }
            }
        }

        //Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
        hyperlinks.get(0).setFieldText( "Modified Text");

        //Save and launch document
        String output = "output/ModifyHyperlinkText.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
