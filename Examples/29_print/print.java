import com.spire.doc.*;

public class print {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/print.docx");
        document.getPrintDocument().print();

    }
}
