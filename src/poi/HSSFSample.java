package poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HSSFSample {
	
	public static final void main(String[] args) {
	HSSFWorkbook wb = new HSSFWorkbook();
		
		HSSFSheet s1 =  wb.createSheet("サンプルシート１");
		
		String[] headerValues = {"列１","列２","合計"};
		int[][] rowValues = {{1,1},{2,2},{3,3},{4,4},{5,5},};
		HSSFRow headerRow = s1.createRow(0);
		
		for (int i=0; i<headerValues.length; i++ ){
			HSSFCell c = headerRow.createCell(i);
			c.setCellValue(headerValues[i]);
		}
		
		for (int i=0; i<rowValues.length; i++ ){
			HSSFRow valueRow = s1.createRow(i+1);
			for (int j=0; j<rowValues[i].length; j++) {
				HSSFCell c = valueRow.createCell(j);
				c.setCellValue(rowValues[i][j]);
			}
			HSSFCell c =valueRow.createCell(2);
			int rowNumber = i+2;
			c.setCellFormula(String.format("A%d*B%d", rowNumber, rowNumber));
		}
		
		
		HSSFSheet s2 =  wb.createSheet("サンプルシート２");
		HSSFSheet s3 =  wb.createSheet("サンプルシート３");
		
		try {
			wb.write(new FileOutputStream(new File("sample1.xls")));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
}
