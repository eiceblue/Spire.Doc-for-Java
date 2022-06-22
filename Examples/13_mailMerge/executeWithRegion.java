import com.spire.doc.*;

public class executeWithRegion {

    public static void main(String[] args) throws Exception {

        String data = "data/dataSample.xml";
        String input = "data/ExecuteMailMergeSample.doc";
        String output = "output/executeWithRegion.doc";

        Document doc = new Document();
        doc.loadFromFile(input);

        //Mail merge
        doc.getMailMerge().executeWidthRegion(data);
        doc.saveToFile(output, FileFormat.Doc);

    }

}
