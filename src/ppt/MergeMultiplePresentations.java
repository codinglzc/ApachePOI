package ppt;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class MergeMultiplePresentations
{

	public static void main(String[] args) throws IOException
	{
		// 创建空的 PPT 对象
		XMLSlideShow ppt = new XMLSlideShow();

		// 准备两个需要合并的 PPT
		String file1 = "C:\\Users\\lc\\Desktop\\hyperlink.pptx";
		String file2 = "C:\\Users\\lc\\Desktop\\TextFormat.pptx";
		String[] inputs = { file1, file2 };

		for (String arg : inputs)
		{

			FileInputStream inputstream = new FileInputStream(arg);
			XMLSlideShow src = new XMLSlideShow(inputstream);

			for (XSLFSlide srcSlide : src.getSlides())
			{
				// 逐个将黄灯片添加到新的 PPT 中
				ppt.createSlide().importContent(srcSlide);
			}
		}

		String file3 = "C:\\Users\\lc\\Desktop\\combinedpresentation.pptx";

		// 创建文件输出流
		FileOutputStream out = new FileOutputStream(file3);

		// 保存
		ppt.write(out);
		System.out.println("合并成功！");
		out.close();
	}

}
