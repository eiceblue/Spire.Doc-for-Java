import com.spire.doc.*;

public class toHtmlExportOption {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/ToHtmlTemplate.docx");
        document.getHtmlExportOptions().setImageEmbedded(true);
        document.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
        document.saveToFile("output/toHtmlExportOption.html", FileFormat.Html);
    }
}
