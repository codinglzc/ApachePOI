package ppt;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PPTtoIMG
{

	public static void main(String[] args) throws IOException
	{
		// 创建现有的 PPT 对象
		File file = new File("C:\\Users\\lc\\Desktop\\combinedpresentation.pptx");
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

		// 获取幻灯片的大小
		Dimension pgsize = ppt.getPageSize();
		
		BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
		
		List<XSLFSlide> slides = ppt.getSlides();
		for (int i = 0; i < slides.size(); i++)
		{
			Graphics2D graphics = img.createGraphics();

			// 初始化画图的范围
			graphics.setPaint(Color.white);
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

			// 渲染
			slides.get(i).draw(graphics);
		}

		// 创建图片输出流
		FileOutputStream out = new FileOutputStream("C:\\Users\\lc\\Desktop\\ppt_image.png");
		javax.imageio.ImageIO.write(img, "png", out);
		ppt.write(out);

		System.out.println("图片生成成功！");
		out.close();
	}

}
