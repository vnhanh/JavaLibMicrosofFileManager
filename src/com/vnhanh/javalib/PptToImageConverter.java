
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class PptToImageConverter {
    void convert(File file) {
        try {
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

            //set zoom factor
            double zoom = .5d;
            AffineTransform at = new AffineTransform();
            at.setToScale(zoom, zoom);

            //get the dimension of size of the slide
            Dimension pgsize = ppt.getPageSize();

            //get slides
            java.util.List<XSLFSlide> slides = ppt.getSlides();

            BufferedImage img = null;
            FileOutputStream out = null;

            for (int i = 0; i < slides.size(); i++) {
                img = new BufferedImage((int)Math.ceil(pgsize.width * zoom), (int)Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();

                //set scaling transform
                graphics.setTransform(at);

                graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

                graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 150);

                //clear the drawing area
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

                //render
                slides.get(i).draw(graphics);

                //creating an image file as output
                out = new FileOutputStream("ppt_image_" + i + ".png");
                ImageIO.write(img, "png", out);
                out.close();
            }
        } catch (Exception e) {
            System.out.println("Caught exception " + e.getMessage());
        }
    }
}
