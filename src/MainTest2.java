import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.net.URL;
import javax.imageio.ImageIO;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class MainTest2 {

	public MainTest2() {
	}

	public static void main(String[] args)throws Exception {
         //����һ��������
          HSSFWorkbook wb=new HSSFWorkbook();
         //����һ�����
          HSSFSheet sheet=wb.createSheet("sheet1");
        //����һ����
         HSSFRow row=sheet.createRow(0);
        //����һ����ʽ
         HSSFCellStyle style=wb.createCellStyle();
       //������Щ��ʽ
       style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
         style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
       //����һ������
        HSSFFont font=wb.createFont();
        font.setColor(HSSFColor.VIOLET.index);
       font.setFontHeightInPoints((short)16);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //������Ӧ�õ���ǰ����ʽ
        style.setFont(font);
       //����һ����ͼ�Ķ���������
       HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
       //��䵥Ԫ��
       for(int i=0;i<5;i++){
           //����һ����Ԫ��
             HSSFCell cell=row.createCell(i);
           switch(i){
               case 0:
               //������ͨ�ı�
                    cell.setCellValue(new HSSFRichTextString("��ͨ�ı�"));
                break;
               case 1:
                  //����Ϊ��״
                    HSSFClientAnchor a1 = new HSSFClientAnchor( 0, 0, 1023, 255, (short) 1, 0, (short) 1, 0 );
                     HSSFSimpleShape shape1 = patriarch.createSimpleShape(a1);
                    //�������������״����ʽ
                    shape1.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
                    
                  break;
               case 2:
                    //����Ϊ������
                    cell.setCellValue(true);
                   break;
                case 3:
                    //����Ϊdoubleֵ
                    cell.setCellValue(12.5);
                   break;
              case 4:
                    //����ΪͼƬ]
                  URL url=MainTest2.class.getResource("baby1.jpg");
                    insertImage(wb,patriarch,getImageData(ImageIO.read(url)),2,4,1);
                    break;
                  
           }
           
          //���õ�Ԫ�����ʽ
            cell.setCellStyle(style);
        }
        FileOutputStream fout=new FileOutputStream("F:\\testImage.xls");
        //������ļ�
       wb.write(fout);
        fout.close();
     }

	// �Զ���ķ���,����ĳ��ͼƬ��ָ��������λ��
	private static void insertImage(HSSFWorkbook wb, HSSFPatriarch pa, byte[] data, int row, int column, int index) {
		int x1 = index * 250;
		int y1 = 0;
		int x2 = x1 + 255;
		int y2 = 255;
		HSSFClientAnchor anchor = new HSSFClientAnchor(x1, y1, x2, y2, (short) column, row, (short) column, row);
		anchor.setAnchorType(2);
		pa.createPicture(anchor, wb.addPicture(data, HSSFWorkbook.PICTURE_TYPE_JPEG));
	}

	// ��ͼƬ����õ��ֽ�����
	private static byte[] getImageData(BufferedImage bi) {
		try {
			ByteArrayOutputStream bout = new ByteArrayOutputStream();
			ImageIO.write(bi, "PNG", bout);
			return bout.toByteArray();
		} catch (Exception exe) {
			exe.printStackTrace();
			return null;
		}
	}
}
