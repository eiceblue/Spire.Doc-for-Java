import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class setGradientBackground {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_2.docx");

        //Set the background type as Gradient.
        document.getBackground().setType(BackgroundType.Gradient);
        BackgroundGradient background = document.getBackground().getGradient();

        //Set the first color and second color for Gradient.
        background.setColor1(Color.white);
        background.setColor2(Color.lightGray);

        //Set the Shading style and Variant for the gradient.
        background.setShadingVariant(GradientShadingVariant.Shading_Down);
        background.setShadingStyle(GradientShadingStyle.Horizontal);

        String result = "output/result-SetGradientBackground.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
