package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class AddImage
{

	public static void main(String[] args) throws IOException
	{
		// 创建一个空的 PPT
	      XMLSlideShow ppt = new XMLSlideShow();
	      
	      // 创建一个幻灯片
	      XSLFSlide slide = ppt.createSlide();
	      
	      // 读一张图片
	      File image=new File("C:\\Users\\lc\\Desktop\\image.jpg");
	      
	      // 将该图片转成一个字节数组
	      byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
	      
	      // 将图片添加到 PPT 中
	      XSLFPictureData idx = ppt.addPicture(picture, PictureType.PNG);
	      
	      // 将图片添加到幻灯片中
	      XSLFPictureShape pic = slide.createPicture(idx);
	      
	      // 创建文件对象和输出流 
	      File file=new File("C:\\Users\\lc\\Desktop\\addingimage.pptx");
	      FileOutputStream out = new FileOutputStream(file);
	      
	      // 保存
	      ppt.write(out);
	      System.out.println("图片添加成功！");
	      out.close();
	}

}
