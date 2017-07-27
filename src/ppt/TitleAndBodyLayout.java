package ppt;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TitleAndBodyLayout
{

	public static void main(String[] args) throws IOException
	{
		// 创建 PPT
		XMLSlideShow ppt = new XMLSlideShow();

		// 获取幻灯片母版
		XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

		// 获取标题和内容布局
		XSLFSlideLayout contentLayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

		// 根据布局创建幻灯片
		XSLFSlide slide = ppt.createSlide(contentLayout);

		// 在幻灯片中选择标题占位符
		XSLFTextShape title1 = slide.getPlaceholder(0);

		// 设置标题
		title1.setText("标题");

		// 在幻灯片中选择内容占位符
		XSLFTextShape body = slide.getPlaceholder(1);

		// 清除该占位符内的文本
		body.clearText();

		// 添加一个段落
		body.addNewTextParagraph().addNewTextRun().setText("这是我的内容。");

		// 创建文件输出流
		File file = new File("C:\\Users\\lc\\Desktop\\contenlayout.pptx");
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		System.out.println("PPT创建成功！");
		out.close();
	}

}
