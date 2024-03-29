import com.spire.doc.*;
import java.util.*;

public class embedNoninstalledFonts {
    public static void main(String[] args) {

        String inputFile="data/convertedTemplate.docx";
        String fontFile="data/PT_Serif-Caption-Web-Regular.ttf";
        String outputFile="output/embedNoninstalledFonts.pdf";

        Document document = new Document();
        document.loadFromFile(inputFile);

        //embed the non-installed fonts.
        ToPdfParameterList parms = new ToPdfParameterList();
        List<PrivateFontPath> fonts = new ArrayList<PrivateFontPath>();
        fonts.add(new PrivateFontPath("PT Serif Caption", fontFile));
        parms.setPrivateFontPaths(fonts);

        //save the document to a PDF file.
        document.saveToFile(outputFile, parms);
    }
}
