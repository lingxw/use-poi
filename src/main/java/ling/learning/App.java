package ling.learning;

import ling.learning.common.ExcelUtils;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	ExcelUtils excel = new ExcelUtils();
    	excel.createXlsx();
    	excel.createSheet("sheet1");
    	
    	Question qa = new Question();
    	qa.setQuestion("12 + 6 = ");
    	qa.setAnswer("18");
    	excel.addData(qa, 0, 0);
    	qa.setQuestion("12 - 6 = ");
    	qa.setAnswer("6");
    	excel.addData(qa, 1, 0);
    	qa.setQuestion("12 - 3 = ");
    	qa.setAnswer("9");
    	excel.addData(qa, 2, 0);
    	excel.addShape(null, 0, 0);
    	excel.save("temp/test.xlsx");
    	
    	excel.createXls();
    	excel.createSheet("sheet1");
    	
    	qa = new Question();
    	qa.setQuestion("12 + 6 = ");
    	qa.setAnswer("18");
    	excel.addData(qa, 0, 0);
    	qa.setQuestion("12 - 6 = ");
    	qa.setAnswer("6");
    	excel.addData(qa, 1, 0);
    	qa.setQuestion("12 - 3 = ");
    	qa.setAnswer("9");
    	excel.addData(qa, 2, 0);
    	excel.addShape(null, 0, 0);
    	excel.save("temp/test.xls");
    }
}
