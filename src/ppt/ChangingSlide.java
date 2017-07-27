package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class ChangingSlide
{
	public static void main(String[] args) throws IOException
	{
		// 创建一个文件对象
		File file = new File("C:\\Users\\lc\\Desktop\\contenlayout.pptx");
		
		// 创建现有的 PPT
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
		
		// 获取当前的页面大小
		java.awt.Dimension pgSize = ppt.getPageSize();
		double pgWidth = pgSize.getWidth();		// 宽
		double pgHeight = pgSize.getHeight();	// 高
		
		System.out.println("当前页面的大小为：");
		System.out.println("width :" + pgWidth);
		System.out.println("height :" + pgHeight);
		
		// 设置一个新的页面大小
		ppt.setPageSize(new java.awt.Dimension(2048, 1536));
		
		// 文件输出流
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		System.out.println("PPT 页面大小更改成功！");
		out.close();
	}
}
