package ling.learning;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

class CreateExcelWithComments {

 static void createCellComment(Cell cell, String commentText) {
  // Create the anchor
  CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
  ClientAnchor anchor = creationHelper.createClientAnchor();
  // When the comment box is visible, have it show in a 1 column x 3 rows space
  anchor.setCol1(cell.getColumnIndex() + 1);
  anchor.setCol2(cell.getColumnIndex() + 2);
  anchor.setRow1(cell.getRow().getRowNum());
  anchor.setRow2(cell.getRow().getRowNum() + 3);
  // Create the comment and set the text
  Drawing drawing = cell.getSheet().createDrawingPatriarch();
  Comment comment = drawing.createCellComment(anchor);
  RichTextString richTextString = creationHelper.createRichTextString(commentText);
  comment.setString(richTextString);
  // Assign the comment to the cell
  cell.setCellComment(comment);
 }

 static void createTriangleShapeTopRight(Cell cell, int width, int height, int r, int g, int b) {
  // Get cell width in pixels
  float columnWidth = cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
  // Get row heigth in pixels
  float rowHeight = cell.getRow().getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI;
  // Create the anchor
  CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
  ClientAnchor anchor = creationHelper.createClientAnchor();
  // Shape starts top, right - shape width and ends top + shape height, right of the cell
  anchor.setCol1(cell.getColumnIndex());
  if (anchor instanceof XSSFClientAnchor) {
   anchor.setDx1(Math.round((columnWidth - width)) * Units.EMU_PER_PIXEL);
  } else if (anchor instanceof HSSFClientAnchor) {
   //see https://stackoverflow.com/questions/48567203/apache-poi-xssfclientanchor-not-positioning-picture-with-respect-to-dx1-dy1-dx/48607117#48607117 for HSSF
   int DEFAULT_COL_WIDTH = 10 * 256;
   anchor.setDx1(Math.round((columnWidth - width) * Units.DEFAULT_CHARACTER_WIDTH / 256f * 14.75f * DEFAULT_COL_WIDTH / columnWidth));
  }
  anchor.setCol2(cell.getColumnIndex() + 1); // left of column index + 1 == right of this cell
  anchor.setDx2(0);
  anchor.setRow1(cell.getRow().getRowNum());
  anchor.setDy1(0);
  anchor.setRow2(cell.getRow().getRowNum());
  if (anchor instanceof XSSFClientAnchor) {
   anchor.setDy2(height * Units.EMU_PER_PIXEL);
  } else if (anchor instanceof HSSFClientAnchor) {
   //see https://stackoverflow.com/questions/48567203/apache-poi-xssfclientanchor-not-positioning-picture-with-respect-to-dx1-dy1-dx/48607117#48607117 for HSSF
   float DEFAULT_ROW_HEIGHT = 12.75f;
   anchor.setDy2(Math.round(height * Units.PIXEL_DPI / Units.POINT_DPI * 14.75f * DEFAULT_ROW_HEIGHT / rowHeight));
  }
  // Create the shape
  Drawing drawing = cell.getSheet().createDrawingPatriarch();
  if (drawing instanceof XSSFDrawing) {
   XSSFSimpleShape shape = ((XSSFDrawing)drawing).createSimpleShape((XSSFClientAnchor)anchor);
   shape.setShapeType(ShapeTypes.RT_TRIANGLE);
   // Flip the shape horizontal and vertical
   shape.getCTShape().getSpPr().getXfrm().setFlipH(true);
   shape.getCTShape().getSpPr().getXfrm().setFlipV(true);
   // Set color
   shape.setFillColor(r, g, b);
  } else if (drawing instanceof HSSFPatriarch) {
   HSSFSimpleShape shape = ((HSSFPatriarch)drawing).createSimpleShape((HSSFClientAnchor)anchor);
   shape.setShapeType(HSSFShapeTypes.RightTriangle);
   // Flip the shape horizontal and vertical
   shape.setFlipHorizontal(true);
   shape.setFlipVertical(true);
   // Set color
   shape.setFillColor(r, g, b);
   shape.setLineStyle(HSSFShape.LINESTYLE_NONE);
  }
 }

 public static void main(String[] args) throws Exception {

  //Workbook workbook = new HSSFWorkbook(); String filePath = "./Excel.xls";
  Workbook workbook = new XSSFWorkbook(); String filePath = "temp/Excel.xlsx";

  Sheet sheet = workbook.createSheet();
  Row row; 
  Cell cell;

  row = sheet.createRow(3);
  cell = row.createCell(5);
  cell.setCellValue("F4");
  sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);   
  createCellComment(cell, "Cell comment for F4");
  createTriangleShapeTopRight(cell, 10, 10, 0, 255, 0);

  row = sheet.createRow(1);
  cell = row.createCell(1);
  cell.setCellValue("B2");
  sheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);   
  createCellComment(cell, "Cell comment for B2");
  createTriangleShapeTopRight(cell, 10, 10, 0, 255, 0);

  try (FileOutputStream out = new FileOutputStream(filePath)) {
   workbook.write(out);
  }

  workbook.close();

 }
}