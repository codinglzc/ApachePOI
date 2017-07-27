package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class DeleteSlide
{

	public static void main(String[] args) throws IOException
	{
		// 打开一个现有的 PPT
		File file = new File("C:\\Users\\lc\\Desktop\\example1.pptx");
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

		// 删除第二个幻灯片
		ppt.removeSlide(1);

		// 创建文件输出流
		FileOutputStream out = new FileOutputStream(file);

		// 保存
		ppt.write(out);
		out.close();
	}

}
