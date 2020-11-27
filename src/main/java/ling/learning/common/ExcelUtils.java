package ling.learning.common;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	private Workbook workbook;
	private Sheet sheet;
	
	public ExcelUtils() {
		this.workbook = null;
		this.sheet = null;
	}
	
	public void createXlsx() {
		this.workbook = new XSSFWorkbook();
		this.sheet = null;
	}
	
//	public void createXlsx2() {
//		this.workbook = new SXSSFWorkbook();
//		this.sheet = null;
//	}
	
	public void createXls() {
		this.workbook = new HSSFWorkbook();
		this.sheet = null;
	}
	
	public boolean createSheet(String name) {
		if (this.workbook == null) {
			return false;
		}
		this.sheet = this.workbook.createSheet(name);
		return true;
	}
	
	public boolean save(String path) {
		if (this.workbook == null) {
			return false;
		}
        FileOutputStream out = null;
		try {
			out = new FileOutputStream(path);
			this.workbook.write(out);
	        out.close();
	        return true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return false;
	}
	
	public boolean addData(Object object, int rowIndex, int colIndex) {
		
		if (object == null) {
			return true;
		}
		
		if (rowIndex < 0) {
			// rowIndex = this.sheet.getLastRowNum();
			rowIndex = 0;
		}
		
		if (colIndex < 0) {
			colIndex = 0;
		}
		
		Row row = this.sheet.createRow(rowIndex);
		Class<?> clazz = object.getClass();
		Field[] declaredFields = clazz.getDeclaredFields();
		int index = -1;
		for(Field field: declaredFields) {
			index++;
		    PropertyDescriptor pd = null;
			try {
				pd = new PropertyDescriptor(field.getName(), clazz);
			    Method getMethod = pd.getReadMethod();
			    Object invoke = getMethod.invoke(object);
				Cell cell=row.createCell(colIndex + index);
				cell.setCellValue(invoke.toString());
			    continue;
			} catch (IntrospectionException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (InvocationTargetException e) {
				e.printStackTrace();
			}
		}
		return false;
	}
	
	public void addShape(Object object, int rowIndex, int colIndex) {
		if (this.workbook instanceof HSSFWorkbook) {
			HSSFCreationHelper helper = ((HSSFWorkbook)this.workbook).getCreationHelper();  
			HSSFPatriarch patriarch = ((HSSFSheet)this.sheet).createDrawingPatriarch();
			  
			//直线  
			HSSFClientAnchor clientAnchor1 = new HSSFClientAnchor(0, 0, 0, 0,  
			        (short) 4, 2, (short) 6, 5);  
			HSSFSimpleShape shape1 = patriarch.createSimpleShape(clientAnchor1);  
			shape1.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);  
			  
			//圆圈（椭圆）  
			HSSFClientAnchor clientAnchor2 = new HSSFClientAnchor(0, 0, 0, 0,  
			        (short) 8, 4, (short) 6, 5);  
			HSSFSimpleShape shape2 = patriarch.createSimpleShape(clientAnchor2);  
			shape2.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);  
			  
			//正方形（长方形）  
			HSSFClientAnchor clientAnchor3 = new HSSFClientAnchor(0, 0, 0, 0,  
			        (short) 12, 6, (short) 6, 5);  
			HSSFSimpleShape shape3 = patriarch.createSimpleShape(clientAnchor3);  
			shape3.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);  
			  
			//Textbox  
			HSSFClientAnchor clientAnchor4 = new HSSFClientAnchor(0, 0, 0, 0,  
			        (short) 14, 8, (short) 6, 5);  
			HSSFTextbox textbox = patriarch.createTextbox(clientAnchor4);  
			textbox.setString(new HSSFRichTextString("This is a test"));  
			//插入图片
			FileInputStream jpeg;
			try {
				jpeg = new FileInputStream("resource/test.png");
				byte[] bytes = IOUtils.toByteArray(jpeg);  
				int pictureIndex = this.workbook.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);  
				jpeg.close();
				HSSFClientAnchor clientAnchor = helper.createClientAnchor();  
				  
				clientAnchor.setCol1(3);  
				clientAnchor.setRow1(14);  
				  
				HSSFPicture picture = patriarch.createPicture(clientAnchor, pictureIndex);  
				picture.resize();  
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}  
		} else if (this.workbook instanceof XSSFWorkbook){
			XSSFCreationHelper helper = ((XSSFWorkbook)this.workbook).getCreationHelper();  
			XSSFDrawing drawing = ((XSSFSheet)this.sheet).createDrawingPatriarch();
			
			//直线  
	        XSSFSimpleShape shape = drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 4, 2, 6, 5));
	        shape.setShapeType(ShapeTypes.LINE);
	        shape.setLineStyleColor(0,0,0);
	        
//	        ClientAnchor anchor = helper.createClientAnchor();
//
//	        anchor.setCol1(0);
//	        anchor.setRow1(0); 
//	        anchor.setCol2(1);
//	        anchor.setRow2(1); 
//	        anchor.setDx1(255);
//	        anchor.setDx2(255);
//	        anchor.setDy1(0);
//	        anchor.setDy2(0);
//
//	        XSSFSimpleShape shape1 = drawing.createSimpleShape((XSSFClientAnchor)anchor);
//	        shape1.setShapeType(ShapeTypes.FLOW_CHART_CONNECTOR);
//	        shape1.setFillColor(255, 0, 0);
			  
			//圆圈（椭圆）  
	        XSSFClientAnchor clientAnchor2 = drawing.createAnchor(0, 0, 0, 0, 6, 4, 8, 5);
	        // clientAnchor2.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE);
			XSSFSimpleShape shape2 = drawing.createSimpleShape(clientAnchor2);  
			shape2.setShapeType(ShapeTypes.ELLIPSE);
			shape2.setFillColor(255, 255, 255);
			shape2.setLineStyleColor(0,0,0);
			//正方形（长方形）  
			XSSFSimpleShape shape3 = drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 6, 5, 12, 6));  
			shape3.setShapeType(ShapeTypes.RECT);  
			shape3.setFillColor(255, 255, 255);
			shape3.setLineStyleColor(0,0,0);
			//Textbox    
			XSSFTextBox textbox = drawing.createTextbox(drawing.createAnchor(0, 0, 0, 0, 6, 5, 14, 8));  
			textbox.setText(new XSSFRichTextString("This is a test")); 
			textbox.setFillColor(255, 255, 255);
			textbox.setLineStyleColor(0,0,0);
			
			//插入图片
			FileInputStream jpeg;
			try {
				jpeg = new FileInputStream("resource/test.png");
				byte[] bytes = IOUtils.toByteArray(jpeg);  
				int pictureIndex = this.workbook.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);  
				jpeg.close();
				XSSFClientAnchor clientAnchor = helper.createClientAnchor();  
				  
				clientAnchor.setCol1(3);  
				clientAnchor.setRow1(14);  
				  
				XSSFPicture picture = drawing.createPicture(clientAnchor, pictureIndex);  
				picture.resize();  
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
