package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class ReadShapes
{

	public static void main(String[] args) throws IOException
	{
		// 创建一个现有的 PPT 对象
		File file = new File("shapes.pptx");
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

		// 获取幻灯片集合
		List<XSLFSlide> slide = ppt.getSlides();

		// 获取形状
		System.out.println("Shapes in the presentation:");
		for (int i = 0; i < slide.size(); i++)
		{
			List<XSLFShape> sh = slide.get(i).getShapes();
			for (int j = 0; j < sh.size(); j++)
			{
				// 输出形状的名称
				System.out.println(sh.get(j).getShapeName());
			}
		}

		FileOutputStream out = new FileOutputStream(file);
		ppt.write(out);
		out.close();
	}

}
