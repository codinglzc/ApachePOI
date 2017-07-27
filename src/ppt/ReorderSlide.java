package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class ReorderSlide
{

	public static void main(String[] args) throws IOException
	{
		// 打开现有的一个 PPT
		File file = new File("C:\\Users\\lc\\Desktop\\example1.pptx");
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

		// 获取幻灯片集合
		List<XSLFSlide> slides = ppt.getSlides();

		// 选择第四个幻灯片
		XSLFSlide selectesdslide = slides.get(3);

		// 把该幻灯片放到第一个
		ppt.setSlideOrder(selectesdslide, 0);

		// 创建文件输出流
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		out.close();
	}

}
