import com.spire.doc.*;

import java.text.DateFormat;
import java.util.Date;

public class setDocumentProperties {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();
        //Load the document from disk.
        document.loadFromFile("data/Sample.docx");

        //Set the build-in Properties.
        document.getBuiltinDocumentProperties().setTitle("Document Demo Document");
        document.getBuiltinDocumentProperties().setAuthor("James");
        document.getBuiltinDocumentProperties().setCompany("e-iceblue");
        document.getBuiltinDocumentProperties().setKeywords("Document, Property, Demo");
        document.getBuiltinDocumentProperties().setComments("This document is just a demo.");

        //Set the custom properties.
        CustomDocumentProperties custom = document.getCustomDocumentProperties();
        custom.add("e-iceblue", true);
        custom.add("Authorized By", "John Smith");
        Date date = new Date();
        DateFormat df = DateFormat.getDateInstance();
        custom.add("Authorized Date", df.format(date));

        String result = "output/setDocumentProperties_out.docx";
        //Save the document.
        document.saveToFile(result, FileFormat.Docx);
    }
}
