import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;
import java.util.ArrayList;

public class removeHyperlinks {
    public static void main(String[] args) {
        //Load Document
        String input =  "data/JAVAHyperlinksTemp_N.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Get all hyperlinks
        ArrayList<Field> hyperlinks = FindAllHyperlinks(doc);

        //Flatten all hyperlinks
        for (int i = hyperlinks.size() - 1; i >= 0; i--)
        {
            FlattenHyperlinks(hyperlinks.get(i));
        }

        //Save and launch document
        String output = "output/RemoveHyperlinks.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
    //Create a method FindAllHyperlinks() to get all the hyperlinks from the sample document
    private static ArrayList<Field> FindAllHyperlinks(Document document)
    {
        ArrayList<Field> hyperlinks = new ArrayList<Field>();

        //Iterate through the items in the sections to find all hyperlinks
        for (Section section : (Iterable<Section>)document.getSections())
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
        return hyperlinks;
    }

    // Flatten the hyperlink field
    private static void FlattenHyperlinks(Field field)
    {
        int ownerParaIndex = field.getOwnerParagraph().ownerTextBody().getChildObjects().indexOf(field.getOwnerParagraph());
        int fieldIndex = field.getOwnerParagraph().getChildObjects().indexOf(field);
        Paragraph sepOwnerPara = field.getSeparator().getOwnerParagraph();
        int sepOwnerParaIndex = field.getSeparator().getOwnerParagraph().ownerTextBody().getChildObjects().indexOf(field.getSeparator().getOwnerParagraph());
        int sepIndex = field.getSeparator().getOwnerParagraph().getChildObjects().indexOf(field.getSeparator());
        int endIndex = field.getEnd().getOwnerParagraph().getChildObjects().indexOf(field.getEnd());
        int endOwnerParaIndex = field.getEnd().getOwnerParagraph().ownerTextBody().getChildObjects().indexOf(field.getEnd().getOwnerParagraph());

        FormatFieldResultText(field.getSeparator().getOwnerParagraph().ownerTextBody(), sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

        field.getEnd().getOwnerParagraph().getChildObjects().removeAt(endIndex);

        for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
        {
            if (i == sepOwnerParaIndex && i == ownerParaIndex)
            {
                for (int j = sepIndex; j >= fieldIndex; j--)
                {
                    field.getOwnerParagraph().getChildObjects().removeAt(j);
                }
            }
            else if (i == ownerParaIndex)
            {
                for (int j = field.getOwnerParagraph().getChildObjects().getCount()-1; j >= fieldIndex; j--)
                {
                    field.getOwnerParagraph().getChildObjects().removeAt(j);
                }

            }
            else if (i == sepOwnerParaIndex)
            {
                for (int j = sepIndex; j >= 0; j--)
                {
                    sepOwnerPara.getChildObjects().removeAt(j);
                }
            }
            else
            {
                field.getOwnerParagraph().ownerTextBody().getChildObjects().removeAt(i);
            }
        }
    }

    //Remove the font color and underline format of the hyperlinks
    private static void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
    {
        for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
        {
            Paragraph para = (Paragraph)ownerBody.getChildObjects().get(i);
            if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
            {
                for (int j = sepIndex + 1; j < endIndex; j++)
                {
                    FormatText((TextRange)para.getChildObjects().get(j));
                }

            }
            else if (i == sepOwnerParaIndex)
            {
                for (int j = sepIndex + 1; j < para.getChildObjects().getCount(); j++)
                {
                    FormatText((TextRange)para.getChildObjects().get(j));
                }
            }
            else if (i == endOwnerParaIndex)
            {
                for (int j = 0; j < endIndex; j++)
                {
                    FormatText((TextRange)para.getChildObjects().get(j));
                }
            }
            else
            {
                for (int j = 0; j < para.getChildObjects().getCount(); j++)
                {
                    FormatText((TextRange)para.getChildObjects().get(j));
                }
            }
        }
    }
    private static void FormatText(TextRange tr)
    {
        //Set the text color to black
        tr.getCharacterFormat().setTextColor(Color.black);
        //Set the text underline style to none
        tr.getCharacterFormat().setUnderlineStyle(UnderlineStyle.None);
    }
}
