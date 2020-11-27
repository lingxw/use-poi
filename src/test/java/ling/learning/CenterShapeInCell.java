package ling.learning;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.util.Units;

import java.io.FileOutputStream;
import java.io.IOException;


class CenterShapeInCell {

 public static void main(String[] args) {
  try {

   Workbook workbook = new XSSFWorkbook();
   Sheet sheet = workbook.createSheet("Sheet1");

   Row row = sheet.createRow(0);
   Cell cell = row.createCell(0);
   row.setHeight((short)(20*20));
   sheet.setColumnWidth(0, 20*256);

   CreationHelper helper = workbook.getCreationHelper();
   Drawing drawing = sheet.createDrawingPatriarch();

   ClientAnchor anchor = helper.createClientAnchor();

   //set anchor to A1 only
   anchor.setCol1(0);
   anchor.setRow1(0); 
   anchor.setCol2(0);
   anchor.setRow2(0); 

   //get the cell width of A1
   float cellWidthPx = sheet.getColumnWidthInPixels(0);
System.out.println(cellWidthPx);

   //set wanted shape size
   int shapeWidthPx = 20;
   int shapeHeightPx = 20;

   //calculate the position of left upper edge
   float centerPosPx = cellWidthPx/2f - (float)shapeWidthPx/2f;
System.out.println(centerPosPx);

   //set the position of left edge as Dx1 in unit EMU
   anchor.setDx1(Math.round(centerPosPx * Units.EMU_PER_PIXEL));

   //set the position of right edge as Dx2 in unit EMU
   anchor.setDx2(Math.round((centerPosPx + shapeWidthPx) * Units.EMU_PER_PIXEL));

   //set upper padding
   int upperPaddingPx = 4;

   //set upper padding as Dy1 in unit EMU
   anchor.setDy1(upperPaddingPx * Units.EMU_PER_PIXEL);

   //set upper padding + shape height as Dy2 in unit EMU
   anchor.setDy2((upperPaddingPx + shapeHeightPx) * Units.EMU_PER_PIXEL);

   XSSFSimpleShape shape = ((XSSFDrawing)drawing).createSimpleShape((XSSFClientAnchor)anchor);
   shape.setShapeType(ShapeTypes.ELLIPSE);
   shape.setFillColor(255, 0, 0);


   FileOutputStream fileOut = new FileOutputStream("temp/CenterShapeInCell.xlsx");
   workbook.write(fileOut);
   fileOut.close();

  } catch (IOException ioex) {
  }
 }
}