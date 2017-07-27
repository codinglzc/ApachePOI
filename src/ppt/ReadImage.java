package ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;

public class ReadImage
{

	public static void main(String[] args) throws IOException
	{
		// 打开现有的 PPT
		File file = new File("C:\\Users\\lc\\Desktop\\addingimage.pptx");
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

		// 读取 PPT 里面所有的图片
		for (XSLFPictureData data : ppt.getPictureData())
		{

			byte[] bytes = data.getData();
			String fileName = data.getFileName();
			PictureType pictureFormat = data.getType();
			System.out.println("picture: " + bytes);
			System.out.println("picture name: " + fileName);
			System.out.println("picture format: " + pictureFormat);
		}

		// saving the changes to a file
		FileOutputStream out = new FileOutputStream(file);
		ppt.write(out);
		out.close();

	}

}
