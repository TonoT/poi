package test;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelT {

	public List<String[]> parseExcel(InputStream inputStream, String suffix, int startrow) throws IOException {
		Workbook workbook = null;
		if ("xls".equals(suffix)) {
			workbook = new HSSFWorkbook(inputStream);
		}
		if (workbook == null) {
			return null;
		}
		Sheet sheet = workbook.getSheetAt(0);
		if (sheet == null) {
			return null;
		}
		int lastRowNum = sheet.getLastRowNum();
		if (lastRowNum <= startrow) {
			return null;
		}
		List<String[]> result = new ArrayList<>();

		Row row = null;
		Cell cell = null;
		for (int rowNum = startrow; rowNum <= lastRowNum; rowNum++) {
			row = sheet.getRow(rowNum);
			short firstCellNum = row.getFirstCellNum();
			short lastCellNum = row.getLastCellNum();
			if (lastCellNum != 0) {
				String[] rowArray = new String[lastCellNum];
				for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
					cell = row.getCell(cellNum);
					if (cell == null) {
						rowArray[cellNum] = null;
					} else {
						rowArray[cellNum] = parseCell(cell);
						
					}
				}
				result.add(rowArray);
			}
		}
		return result;
	}

	public String parseCell(Cell cell) {
		String cellStr = null;
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			cellStr = cell.getRichStringCellValue().toString();
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			cellStr = "";
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf=null;
				if (cell.getCellType()==HSSFDataFormat.getBuiltinFormat("h:mm")) {
					sdf = new SimpleDateFormat("HH:mm");
				}else{
					sdf = new SimpleDateFormat("yyyy-MM-dd");
				}
				Date temp = cell.getDateCellValue();
				cellStr = sdf.format(temp);
			}else{
				double temp = cell.getNumericCellValue();
				cellStr = ""+temp;
			}

			break;
		default:
			break;

		}
		return cellStr;
	}
}
