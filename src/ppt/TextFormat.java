package ppt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TextFormat
{

	public static void main(String[] args) throws IOException
	{
		// 创建空的 PPT 对象
		XMLSlideShow ppt = new XMLSlideShow();

		// 获取 PPT 母版
		XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

		// 选择标题和内容布局
		XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

		// 根据标题创建幻灯片
		XSLFSlide slide = ppt.createSlide(slidelayout);

		// 获取该幻灯片内容对象
		XSLFTextShape body = slide.getPlaceholder(1);

		// 清空内容
		body.clearText();

		// 添加一个段落
		XSLFTextParagraph paragraph = body.addNewTextParagraph();

		// 第一行
		XSLFTextRun run1 = paragraph.addNewTextRun();
		
		// 设置内容的 Text
		run1.setText("第一行");

		// 设置字体颜色
		run1.setFontColor(java.awt.Color.red);

		// 设置字体大小
		run1.setFontSize(24.0);

		// 添加换行符
		paragraph.addLineBreak();

		// 第二行
		XSLFTextRun run2 = paragraph.addNewTextRun();
		run2.setText("第二行");
		run2.setFontColor(java.awt.Color.CYAN);

		// 设置字体加粗
		run2.setBold(true);
		paragraph.addLineBreak();

		// 第三行
		XSLFTextRun run3 = paragraph.addNewTextRun();
		run3.setText(" 第三行");
		run3.setFontSize(12.0);

		// 设置字体 italic
		run3.setItalic(true);

		// strike through the text
		run3.setStrikethrough(true);
		paragraph.addLineBreak();

		// 第四行
		XSLFTextRun run4 = paragraph.addNewTextRun();
		run4.setText(" 第四行");
		// 下划线
		run4.setUnderlined(true);
		paragraph.addLineBreak();

		// 创建文件输出流
		File file = new File("C:\\Users\\lc\\Desktop\\TextFormat.pptx");
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		out.close();
	}

}
