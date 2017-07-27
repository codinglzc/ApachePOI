package ppt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;

public class SlideLayouts
{

	public static void main(String[] args)
	{
		// 创建一个空的 PPT
		XMLSlideShow ppt = new XMLSlideShow();
		System.out.println("可用的幻灯片布局");
		
		// 获取幻灯片母版列表
		for (XSLFSlideMaster master : ppt.getSlideMasters())
		{
			// 获取幻灯片每个母版的布局
			for (XSLFSlideLayout layout : master.getSlideLayouts())
			{
				// 获取幻灯片布局的类型名称
				System.out.println(layout.getType());
			}
		}
	}

}
