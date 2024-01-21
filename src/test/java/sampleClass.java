import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;  
import java.text.SimpleDateFormat;  
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.concurrent.TimeUnit;
import java.util.ArrayList;
import java.util.Collections;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class sampleClass {
    public static boolean findDiff(long startDate,long endDate) {
    		
   		    long timeDiff=startDate-endDate;
    		
    		long hrsDiff=TimeUnit.MILLISECONDS.toHours(timeDiff)%24;
//      		System.out.println(hrsDiff);
    		
    		if(hrsDiff<10 && hrsDiff>1) {
   			    return true;
    		}
         	return false;
    }
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String excelFilePath=System.getProperty("user.dir")+"\\files\\Assignment_Timecard.xlsx";
		File excelFile=new File(excelFilePath);
        FileInputStream fis=new FileInputStream(excelFile);
        
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        
        BufferedWriter writer=new BufferedWriter(new FileWriter("output.txt"));
         
        int rowsLength=sheet.getLastRowNum();
        
        writer.write("\nName of the Employees and the Position who worked for 7 consecutive days->");
        writer.write("\n");
        writer.write("\nEmployee Name\t\t\t\t\t\tPosition");
        writer.write("\n");
        int count=1;
        for(int r=2;r<=rowsLength;r++) {
        	XSSFRow row1= sheet.getRow(r);
        	XSSFRow rowp= sheet.getRow(r-1);
        	XSSFCell cell=row1.getCell(0);
        	XSSFCell cellp=rowp.getCell(0);
        	if(cell.getStringCellValue()==cellp.getStringCellValue()) {
        		count++;
        	}
        	else {
        		if(count>=7) {
        			writer.write("\n");
        			writer.write(rowp.getCell(7).getStringCellValue()+"\t\t\t\t\t"+cellp.getStringCellValue());
        		}
        		count=1;
        	}
        }
        
        
        
        
       writer.write("\nName Of The Employees and Their Position who have less than 10 hours of time between shifts but greater then 1 hour->");
       writer.write("\n");
       writer.write("\nEmployee Name\t\t\t\t\tPosition");
       writer.write("\n");
       
       for(int r=2;r<=rowsLength;r++) {
        	XSSFRow rowc=sheet.getRow(r);
        	XSSFCell cellc=rowc.getCell(2);
        	XSSFRow rowp=sheet.getRow(r-1);
        	XSSFCell cellp=rowp.getCell(3);
        	
        	HashSet<String>hs= new HashSet<>();
        	if(cellp.getCellType()==CellType.NUMERIC && cellc.getCellType()==CellType.NUMERIC) {
        		Date date = cellc.getDateCellValue();
        		Date date1=cellp.getDateCellValue();
        		if(rowc.getCell(7).getStringCellValue()==rowp.getCell(7).getStringCellValue() && findDiff(date.getTime(),date1.getTime()) && !hs.contains(rowp.getCell(7).getStringCellValue())) {
        			    writer.write("\n");
            			writer.write(rowp.getCell(7).getStringCellValue()+"\t\t\t\t\t\t"+rowp.getCell(0).getStringCellValue());
                        hs.add(rowp.getCell(7).getStringCellValue());
        		}
        	}        	
        }
       
       
       writer.write("\nName Of The Employees who worked for 14 hours in a single shift->");
       writer.write("\n");
       writer.write("\nEmployee Name\t\t\t\t\t\t\tPosition");
       writer.write("\n");
       DataFormatter dataFormatter = new DataFormatter();
       for(int r=1;r<rowsLength;r++) {
    	   String timeString = dataFormatter.formatCellValue(sheet.getRow(r).getCell(4));
    	   if(!timeString.isEmpty()) {
    		   String[] timeParts = timeString.split(":");
    		   int hours = Integer.parseInt(timeParts[0]);
               if(hours>14) {
            	   writer.write("\n");
            	   writer.write(sheet.getRow(r).getCell(7).getStringCellValue()+"\t\t\t\t"+sheet.getRow(r).getCell(0).getStringCellValue());
               }
    	   }
    	   
       }
       writer.close();
	}
}