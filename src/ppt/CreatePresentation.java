package ppt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class CreatePresentation
{

	public static void main(String[] args) throws IOException
	{
		// 创建空 PPT
		XMLSlideShow ppt = new XMLSlideShow();
		
		// 创建一个 FileOutputStream 对象
		File file = new File("C:\\Users\\lc\\Desktop\\example1.pptx");
		FileOutputStream out = new FileOutputStream(file);
		
		// 保存更改到这个文件中
		ppt.write(out);
		System.out.println("PPT 创建成功！");
		out.close();
		
	}

}
