import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class setTextWrap {
    public static void main(String[] args) {
        String input="data/imageTemplate.docx";
        String output="output/setTextWrap.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        for (int i=0;i<doc.getSections().getCount();i++)
        {
            Section sec =doc.getSections().get(i);
            for (int j=0;j<sec.getParagraphs().getCount();j++)
            {
                Paragraph para= sec.getParagraphs().get(j);
                //get all pictures in the Word document
                for (int p=0;p<para.getChildObjects().getCount();p++)
                {
                    DocumentObject docObj = para.getChildObjects().get(p);
                    if (docObj.getDocumentObjectType() == DocumentObjectType.Picture)
                    {
                        //set text wrap styles for each piture
                        DocPicture picture = (DocPicture)docObj;
                        picture.setTextWrappingStyle( TextWrappingStyle.Through);
                        picture.setTextWrappingType( TextWrappingType.Both);
                    }
                }
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
