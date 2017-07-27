package word;

import java.io.FileInputStream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordExtractor
{

	public static void main(String[] args) throws Exception
	{
		XWPFDocument docx = new XWPFDocument(new FileInputStream("create_paragraph.docx"));
		// using XWPFWordExtractor Class
		XWPFWordExtractor we = new XWPFWordExtractor(docx);
		System.out.println(we.getText());
	}

}
