import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.io.*;

public class extractOLE {
    public static void main(String[] args) {
        //Create document and load file from disk
        Document doc = new Document();
        doc.loadFromFile("data/extractOLE.docx");

        //Traverse through all sections of the word document
        for (int s = 0; s < doc.getSections().getCount(); s++) {
            Section section = doc.getSections().get(s);
            //Traverse through all Child Objects in the body of each section
            for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
                DocumentObject obj = section.getBody().getChildObjects().get(i);
                //Find the paragraph
                if (obj instanceof Paragraph) {
                    Paragraph par = (Paragraph) obj;
                    for (int j = 0; j < par.getChildObjects().getCount(); j++) {
                        DocumentObject o = par.getChildObjects().get(j);
                        //Check whether the object is OLE
                        if (o.getDocumentObjectType() == DocumentObjectType.Ole_Object) {
                            DocOleObject ole = (DocOleObject) o;
                            String type = ole.getObjectType();

                            //Check whether the object type is "Acrobat.Document.11"
                            if ("AcroExch.Document.DC".equals(type)) {
                                //Write the data of OLE into file
                                byteArrayToFile(ole.getNativeData(), "output/extractOLE.pdf");
                            }

                            //Check whether the object type is "Excel.Sheet.8"
                            else if ("Excel.Sheet.8".equals(type)) {
                                byteArrayToFile(ole.getNativeData(), "output/extractOLE.xls");
                            }

                            //Check whether the object type is "PowerPoint.Show.12"
                            else if ("PowerPoint.Show.12".equals(type)) {
                                byteArrayToFile(ole.getNativeData(), "output/extractOLE.pptx");
                            }
                        }
                    }
                }
            }
        }
    }

    public static void byteArrayToFile(byte[] datas, String destPath) {
        File dest = new File(destPath);
        try (InputStream is = new ByteArrayInputStream(datas);
             OutputStream os = new BufferedOutputStream(new FileOutputStream(dest, false));) {
            byte[] flush = new byte[1024];
            int len = -1;
            while ((len = is.read(flush)) != -1) {
                os.write(flush, 0, len);
            }
            os.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
