
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
public static String sFile = "./data/New.xlsx";
public static FileInputStream oRead;
public static FileOutputStream oWrite;
public static String sSheet = "Course";
public static void main(String[] args) throws Exception {
	getCellValueBasedonRowNCell(sSheet, 1, 2);
	//getAllCellValues(sSheet);
	//getAllCellValuesAndWrite(sSheet);
}

public static void getCellValueBasedonRowNCell(String sSheet,int row,int col) throws Exception {
	oRead = new FileInputStream(sFile);
	System.out.println(sFile);
	// When excel extension is .xlsx then use XSSF class
	// When excel extension is .xls then use HSSF class	
	XSSFWorkbook wb =  new XSSFWorkbook(oRead);
	XSSFSheet oSheet = wb.getSheet(sSheet);
	XSSFRow oRow = oSheet.getRow(row);
    XSSFCell oCell = oRow.getCell(col);
	String value = oCell.getStringCellValue();
	System.out.println("Value in "+row+" and "+col+" is : "+value);
	wb.close();
	oRead.close();
}
public static void getAllCellValues(String sSheet) throws Exception {
	 oRead = new FileInputStream(sFile);
	XSSFWorkbook wb = new XSSFWorkbook(oRead);
	XSSFSheet oSheet = wb.getSheet(sSheet);
	XSSFRow oRow;
	XSSFCell oCell;
	int lastRowNum = oSheet.getLastRowNum();
	System.out.println("Last Row Number is : "+lastRowNum);
	for(int iRow=0;iRow<=lastRowNum;iRow++) {
		oRow = oSheet.getRow(iRow);
		short lastCellNum = oRow.getLastCellNum();
		for(int iCell=0;iCell<lastCellNum;iCell++) {
			oCell = oRow.getCell(iCell);
			CellType cellType = oCell.getCellType();
			switch (cellType) {
			case NUMERIC:
				System.out.print(oCell.getNumericCellValue()+"	");
				break;
			case STRING:
				System.out.print(oCell.getStringCellValue()+"	");
				break;
			case BOOLEAN:
				System.out.print(oCell.getBooleanCellValue()+"	");
				break;
			default:
				System.out.print("Cell Type is not valid");
				break;
			}
		}
		System.out.println("");
	}
	wb.close();
	oRead.close();
}

public static void getAllCellValuesAndWrite(String sSheet) throws Exception {
	oRead = new FileInputStream(sFile);
	XSSFWorkbook wb = new XSSFWorkbook(oRead);
	XSSFSheet oSheet = wb.getSheet(sSheet);
	XSSFRow oRow;
	XSSFCell oCell;
	int lastRowNum = oSheet.getLastRowNum();
	System.out.println("Last Row Number is : "+lastRowNum);
	for(int iRow=0;iRow<=lastRowNum;iRow++) {
		oRow = oSheet.getRow(iRow);
		short lastCellNum = oRow.getLastCellNum();
		for(int iCell=0;iCell<lastCellNum;iCell++) {
			oCell = oRow.getCell(iCell);
			CellType cellType = oCell.getCellType();
			switch (cellType) {
			case NUMERIC:
				System.out.print(oCell.getNumericCellValue()+"	");
				break;
			case STRING:
				System.out.print(oCell.getStringCellValue()+"	");
				break;
			case BOOLEAN:
				System.out.print(oCell.getBooleanCellValue()+"	");
				break;
			default:
				System.out.print("Cell Type is not valid");
				break;
			}
			if(iRow!=0) {
				oSheet.getRow(iRow).createCell(4).setCellValue("Joined");
				oWrite = new FileOutputStream(sFile);
				wb.write(oWrite);
				oWrite.close();
			}
		}
		System.out.println("");
	}
	
	wb.close();
	oRead.close();
}		
}



