package it.l_soft.EmailDispatcher;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class ReadFromExcel {
	
	POIFSFileSystem fs;
    OutputStream os;
	HSSFWorkbook wb;
	HSSFSheet sheet;
    HSSFRow row;
    HSSFCell cell;
	String[] emailCCValue;
	String emailBCCValue;
	String fileAttachValue;
	String excelEmailAttachFilesExt;
	String excelEmailAttachFilesFolder;
	String excelFilePath;
	String cellValue;
	DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
 
	int rows; // No of rows
    int rowIdx; // current row
    int cols = 0; // No of columns
    int emailColIdx = -1, emailBCCColIdx = -1, fileAttachColIdx = -1, emailCCNo = -1;
    int[] emailCCColIdx = new int[10];

	private Map<String, Integer> columnList = new HashMap<>();

	public ReadFromExcel(int excelHeaderRow, int startFrom, ApplicationProperties ap) throws IOException
	{
		this.excelFilePath = ap.getExcelFilePath();
	    fs = new POIFSFileSystem(new FileInputStream(excelFilePath));
	    wb = new HSSFWorkbook(fs);
	    sheet = wb.getSheet(ap.getSheetName());
	    rows = sheet.getPhysicalNumberOfRows();
	    String[] emailCCList = ap.getExcelEmailCCFieldName().split(";");

	    this.excelEmailAttachFilesFolder = ap.getExcelEmailAttachFilesFolder();
	    this.excelEmailAttachFilesExt = ap.getExcelEmailAttachFilesExt();
	    
    	row = sheet.getRow(excelHeaderRow);
	    for(int col = 0; (col < row.getLastCellNum()); col++)
	    {
	    	cell = row.getCell(col);
	    	cellValue = formatter.formatCellValue(cell);
	    	columnList.put(cellValue, col);
	    	
	    	try {
		    	if (cellValue == null)
		    		break;
		    	if (cellValue.compareTo(ap.getExcelEmailTOFieldName()) == 0) // cell.getStringCellValue()
		    	{
		    		emailColIdx = col;
		    	}
		    	else if ((cellValue.compareTo(ap.getExcelEmailBCCFieldName()) == 0) && (ap.getExcelEmailBCCFieldName().compareTo("") != 0))
		    	{
		    		emailBCCColIdx = col;
		    	}
		    	else if ((cellValue.compareTo(ap.getExcelEmailAttachFieldName()) == 0) && (ap.getExcelEmailAttachFieldName().compareTo("") != 0))
		    	{
		    		fileAttachColIdx = col;
		    	}
		    	else
		    	{
		    		for(int y = 0; y < emailCCList.length; y++)
		    		{
			    		if ((cellValue.compareTo(emailCCList[y]) == 0) && (ap.getExcelEmailCCFieldName().compareTo("") != 0))
			    		{
			    			emailCCColIdx[++emailCCNo] = col;
			    			break;
			    		}
		    		}
		    	}
	    	} 
	    	catch(Exception e)
	    	{
	    		e.printStackTrace();
	    	}
	    }
	    if (emailCCNo == -1)
	    {
	    	emailCCValue = new String[1];
	    	emailCCValue[0] = ap.getExcelEmailCCFieldName();
	    }
	    else
	    {
	    	emailCCValue = new String[++emailCCNo];
	    }
	    if (emailBCCColIdx == -1)
	    	emailBCCValue = ap.getExcelEmailBCCFieldName();
	    if (fileAttachColIdx == -1)
	    	fileAttachValue = ap.getExcelEmailTOFieldName();

	    rowIdx = startFrom;
	}
	
	public HSSFRow getNextRow()
	{
	    if (rowIdx <= rows)
	    {
	    	row = sheet.getRow(rowIdx++);
	    }
	    else
	    {
	    	row = null;
	    }
	    return row;
	}
	
	public String getEmail()
	{
		String addresses = getField(emailColIdx).replaceAll(";", ",");
		String[] addrArray = addresses.split(","); 
		for(String address : addrArray)
		{
			try {
				InternetAddress emailAddr = new InternetAddress(address);
				emailAddr.validate();
			} 
			catch (AddressException ex) {
				return "";
			}
		}
		return addresses;
	}

	public String getEmailCCValue() 
	{
    	String retVal = "";
	    if (emailCCNo == -1)
			retVal = emailCCValue[0];
	    else
	    {
	    	String sep = "";
	    	for(int i = 0; i < emailCCNo; i++)
	    	{
	    		if ((getField(emailCCColIdx[i]) == null) || (getField(emailCCColIdx[i]).trim().compareTo("") == 0))
	    			continue;
	    		retVal += sep + getField(emailCCColIdx[i]);
	    		sep = ",";
	    	}
	    	retVal = retVal.replaceAll(";", ",");
	    }
	    return retVal;
	}

	public String getEmailBCCValue() {
	    if (emailBCCColIdx == -1)
	    	return emailBCCValue;
	    else
	    	return getField(emailBCCColIdx).replaceAll(";", ",");
	}

	public String getFileAttachValue() {
	    if (fileAttachColIdx == -1)
			return fileAttachValue;
	    else
			return excelEmailAttachFilesFolder + getField(fileAttachColIdx)  + excelEmailAttachFilesExt;
	}

	public String getField(int idx)
	{
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

		if (row.getCell(idx) == null)
    	{
    		cellValue = "";
    	}
    	else
    	{
    	    if(row.getCell(idx).getCellType() == Cell.CELL_TYPE_FORMULA) 
    	    {
    	    	 cellValue = evaluator.evaluate(row.getCell(idx)).formatAsString().replaceAll("\"",  "").trim();
    	    }
    	    else
    	    {
    	    	cellValue = formatter.formatCellValue(row.getCell(idx));
    	    }
    	}
		return(cellValue);
	}
	
	public String getField(String fieldName)
	{
		Integer idx = columnList.get(fieldName);
		cellValue = "";
		if (idx != null)
		{
    		cellValue = getField(idx);
		}
		return(cellValue);
	}

	public void setSentFlag(int colIdx, String value)
	{
		cell = row.createCell(colIdx);
		if (cell != null)
			cell.setCellValue(value);
	}
	
	public void writeChanges() throws IOException
	{
		String fileNameCopy = excelFilePath.replace(".xls", "-copy.xls");
	    os = new FileOutputStream(fileNameCopy);
		wb.write(os);
		os.close();
	}

}
