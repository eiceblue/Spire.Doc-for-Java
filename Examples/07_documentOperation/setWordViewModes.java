import com.spire.doc.*;
public class setWordViewModes {
    public static void main(String[] args){
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Set Word view modes.
        document.getViewSetup().setDocumentViewType(DocumentViewType.Web_Layout);
        document.getViewSetup().setZoomPercent(150);
        document.getViewSetup().setZoomType(ZoomType.None);

        String result = "output/setWordViewModes_out.docx";
        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
