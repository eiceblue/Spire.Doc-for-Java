import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.Field;
import java.io.*;
import java.util.ArrayList;

public class findHyperlinks {
    public static void main(String[] args) throws IOException {
        //Load Document
        String input = "data/JAVAHyperlinksTemp_N.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Create a hyperlink list
        ArrayList<Field> hyperlinks = new ArrayList<Field>();
        String hyperlinksText = "";
        //Iterate through the items in the sections to find all hyperlinks
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
                                //Get the hyperlink text
                                hyperlinksText += field.getFieldText() + "\r\n";
                            }
                        }
                    }
                }
            }
        }

        //Save the text of all hyperlinks to TXT File and launch it
        String output = "output/HyperlinksText.txt";
        writeStringToTxt(hyperlinksText,output);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        File file=new File(txtFileName);
        if (file.exists())
        {
            file.delete();
        }
        FileWriter fWriter = new FileWriter(txtFileName, true);
        try {
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
