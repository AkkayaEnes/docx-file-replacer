import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocumentUpdater {
    public static void updateDocument(String sourceFilePath, String destinationFilePath, Map<String, String> placeholders, String fontName, int fontSize) {
        try {
            FileInputStream fileInputStream = new FileInputStream(sourceFilePath);
            XWPFDocument sourceDocument = new XWPFDocument(fileInputStream);
            XWPFDocument destinationDocument = new XWPFDocument();

            for (XWPFParagraph sourceParagraph : sourceDocument.getParagraphs()) {
                XWPFParagraph destinationParagraph = destinationDocument.createParagraph();

                for (XWPFRun sourceRun : sourceParagraph.getRuns()) {
                    String text = sourceRun.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                            String placeholder = "{" + entry.getKey() + "}";
                            if (text.contains(placeholder)) {
                                text = text.replace(placeholder, entry.getValue());
                            }
                        }
                        XWPFRun destinationRun = destinationParagraph.createRun();

                        //f
                        destinationRun.setFontFamily(fontName);
                        destinationRun.setFontSize(fontSize);

                        destinationRun.setText(text);
                    }
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(destinationFilePath);
            destinationDocument.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Map<String, String> placeholders = new HashMap<>();
        placeholders.put("id", "12345");
        placeholders.put("name", "John Doe");
        placeholders.put("age", "30");
        placeholders.put("gender", "Male");

        String sourceFilePath = "C:/Users/Enes/Desktop/JavaProjeleri/regexanddatastruc/exam.docx";
        String destinationFilePath = "C:/Users/Enes/Desktop/JavaProjeleri/regexanddatastruc/examss.docx";


        String fontName = "Arial";
        int fontSize = 12;

        updateDocument(sourceFilePath, destinationFilePath, placeholders, fontName, fontSize);
    }
}