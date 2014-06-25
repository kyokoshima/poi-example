package poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hslf.model.Picture;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextBox;
import org.apache.poi.hslf.usermodel.SlideShow;

public class HSLFSample {

	public static void main(String[] args) throws Exception {
		SlideShow ss = new SlideShow();
		
		Slide s1 = ss.createSlide();
		TextBox title = s1.addTitle();
		title.setText("ÉTÉìÉvÉãï∂èë");
		
		int picIndex = ss.addPicture(new File("project-logo.jpg"), Picture.JPEG);
		Picture logo = new Picture(picIndex);
		
		s1.addShape(logo);
		
		ss.write(new FileOutputStream(new File("sample.ppt")));
	}

}
