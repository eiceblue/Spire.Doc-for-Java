import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

/**
 * Remove the editable range of the document.
 */
public class removeEditableRange {
    public static void main(String []args){
        String inputFile="data/removeEditableRange.docx";
        String outputFile="removeEditableRange_output.docx";
        //Create and load the Word file
        Document document=new Document();
        document.loadFromFile(inputFile);
        //Find "PermissionStart" and "PermissionEnd" tags and remove them
        for(int j=0;j<document.getSections().getCount();j++){
            Section section=document.getSections().get(j);
            for(int k=0;k<section.getParagraphs().getCount();k++){
                Paragraph paragraph=section.getParagraphs().get(k);
                for(int i=0;i<paragraph.getChildObjects().getCount();){
                    DocumentObject obj=paragraph.getChildObjects().get(i);
                    if(obj instanceof PermissionStart||obj instanceof PermissionEnd){
                        paragraph.getChildObjects().remove(obj);
                    }else{
                        i++;
                    }
                }
            }
        }
        //Save the document
        document.saveToFile(outputFile,FileFormat.Docx);
    }
}
