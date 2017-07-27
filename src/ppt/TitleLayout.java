package ppt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TitleLayout
{

	public static void main(String[] args) throws IOException
	{
		// 创建 PPT
		XMLSlideShow ppt = new XMLSlideShow();
		
		// 获取幻灯片母版
		XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
		
		// 获取标题布局
		XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
		
		// 根据布局创建幻灯片
		XSLFSlide slide = ppt.createSlide(titleLayout);
		
		// 在幻灯片中选择一个占位符
		XSLFTextShape title1 = slide.getPlaceholder(0);
		
		// 设置标题
		title1.setText("标题");
		
		// 创建文件输出流
		File file = new File("C:\\Users\\lc\\Desktop\\Titlelayout.pptx");
		FileOutputStream out = new FileOutputStream(file);
		
		// 保存
		ppt.write(out);
		System.out.println("PPT 创建成功！");
		out.close();
	}

}
