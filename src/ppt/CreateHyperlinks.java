package ppt;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class CreateHyperlinks
{

	public static void main(String[] args) throws IOException
	{
		// 创建空的 PPT
		XMLSlideShow ppt = new XMLSlideShow();

		// 获取 PPT 的母版集合中的第一个
		XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

		// 创建标题和内容的布局
		XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

		// 根据布局创建幻灯片
		XSLFSlide slide = ppt.createSlide(slidelayout);

		// 获取该幻灯片的内容对象
		XSLFTextShape body = slide.getPlaceholder(1);

		// 清除内容
		body.clearText();

		// 为内容添加一个段落
		XSLFTextRun textRun = body.addNewTextParagraph().addNewTextRun();

		// 设置内容的文字
		textRun.setText("内容");

		// 创建超链接
		XSLFHyperlink link = textRun.createHyperlink();

		// 为超链接设置链接地址
		link.setAddress("http://www.baidu.com/");

		// 创建文件对象和文件输出流
		File file = new File("C:\\Users\\lc\\Desktop\\hyperlink.pptx");
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		System.out.println("创建 PPT 成功！");
		out.close();
	}

}
