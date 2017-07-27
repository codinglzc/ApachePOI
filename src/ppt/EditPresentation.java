package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class EditPresentation
{

	public static void main(String[] args) throws IOException
	{
		// 打开一个现有的 PPT
		File file = new File("C:\\Users\\lc\\Desktop\\example1.pptx");
		FileInputStream in = new FileInputStream(file);
		XMLSlideShow ppt = new XMLSlideShow(in);
		
		// 向 PPT 添加幻灯片
		XSLFSlide slide1 = ppt.createSlide();
		XSLFSlide slide2 = ppt.createSlide();
		
		// 保存更改
		FileOutputStream out = new FileOutputStream(file);
		ppt.write(out);
		
		System.out.println("PPT 编辑成功！"); 
		out.close();
	}
}
