import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class createTableOfContentByDefault {
    public static void main(String[] args) {
        Document doc = new Document();
        Section section = doc.addSection();
        Paragraph para = section.addParagraph();
        //Create table of content with default switches(\o "1-3" \h \z)
        para.appendTOC(1, 3);

        Paragraph par = section.addParagraph();
        TextRange tr = par.appendText("Flowers");
        tr.getCharacterFormat().setFontSize(30);
        par.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Create paragraph and set the head level
        Paragraph para1 = section.addParagraph();
        para1.appendText("Ornithogalum");
        //Apply the Heading1 style
        para1.applyStyle(BuiltinStyle.Heading_1);
        //Add paragraphs
        para1 = section.addParagraph();
        DocPicture picture = para1.appendPicture("data/ornithogalum.jpg");
        picture.setTextWrappingStyle(TextWrappingStyle.Square);
        para1.appendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.");
        para1 = section.addParagraph();


        Paragraph para2 = section.addParagraph();
        para2.appendText("Rosa");
        //Apply the Heading2 style
        para2.applyStyle(BuiltinStyle.Heading_2);
        para2 = section.addParagraph();
        DocPicture picture2 = para2.appendPicture("data/rosa.jpg");
        picture2.setTextWrappingStyle(TextWrappingStyle.Square);
        para2.appendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.");
        section.addParagraph();

        Paragraph para3 = section.addParagraph();
        para3.appendText("Hyacinth");
        //Apply the Heading3 style
        para3.applyStyle(BuiltinStyle.Heading_3);
        para3 = section.addParagraph();
        DocPicture picture3 = para3.appendPicture("data/hyacinths.JPG");
        picture3.setTextWrappingStyle(TextWrappingStyle.Tight);
        para3.appendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae.These are commonly called hyacinths.The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).");
        para3 = section.addParagraph();
        para3.appendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.");

        //Update TOC
        doc.updateTableOfContents();
        //Save to file
        String output = "output/createTableOfContentByDefault.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
