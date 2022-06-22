import com.spire.doc.*;
import com.spire.doc.fields.omath.*;

public class GetMathEquation {
    public static void main(String[] args) {
        Document doc = new Document();
        doc.loadFromFile("data/GetMathEquation.docx");
        StringBuilder stringBuilder = new StringBuilder();
        for(int i=0;i<doc.getSections().getCount();i++){
            for(int j=0;j<doc.getSections().get(i).getParagraphs().getCount();j++){
                for(int k=0;k<doc.getSections().get(i).getParagraphs().get(j).getChildObjects().getCount();k++){
                    DocumentObject obj =doc.getSections().get(i).getParagraphs().get(j).getChildObjects().get(k);
                    if (obj instanceof OfficeMath) {
                        OfficeMath math=(OfficeMath)doc.getSections().get(i).getParagraphs().get(j).getChildObjects().get(k);
                        stringBuilder.append(math.toMathMLCode()+"\n");
                    }
                }
            }
        }
        System.out.println(stringBuilder);
    }
}
