package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Function;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import model.LoadingOFC2;

public class ExcelReader {
	
	private static final int VVD_CODE = 0;
	private static final int SERVICE_LANE_CODE = 2;
	private static final int SALES_WEEK = 3;
	private static final int SALES_MAIN_OFFICE_CODE = 4;
	private static final int POL_NODE_CODE = 5;
	private static final int POD_NODE_CODE = 8;
	private static final int OCN_IPC = 9;
	private static final int BKG_TEU_ALL = 16;
	private static final int BKG_WGT_ALL = 17;
	private static final int BKG_TEU_FIRM = 18;
	private static final int BKG_WGT_FIRM = 19;
	
	private static final String PATH = "C:\\test-data\\test-data.xlsx";

	
	public static Map<String, List<LoadingOFC2>> readDataFromSheet(String tabName){
		Map<String, List<LoadingOFC2>> result = new HashMap<String, List<LoadingOFC2>>();
		
		List<LoadingOFC2> lstData = new ArrayList<LoadingOFC2>();
		try {
			// Đọc Excel lưu vào List<Object> lstData
			InputStream inputStream = new FileInputStream(new File(PATH));
			
			Workbook workbook = getWorkbook(inputStream, PATH);
			
			Sheet sheet = workbook.getSheet(tabName);
			Iterator<Row> iterator = sheet.iterator();

			while(iterator.hasNext()) {
				Row currentRow = iterator.next();
				// Check if current row is header => skip to next row
				if(currentRow.getRowNum() == 0) {
					continue;
				}
				
				Iterator<Cell> cells = currentRow.cellIterator();
				
				LoadingOFC2 loadingOfc = new LoadingOFC2();
				while(cells.hasNext()) {
					Cell currentCell = cells.next();
					
					int index = currentCell.getColumnIndex();
					
					switch (index) {
						case VVD_CODE:
							loadingOfc.setVvdCode(currentCell.getStringCellValue());
							break;
						case SERVICE_LANE_CODE:
							loadingOfc.setServiceLandCode(currentCell.getStringCellValue());
							break;
						case SALES_WEEK:
							loadingOfc.setSalesWeek(currentCell.getStringCellValue());
							break;
						case SALES_MAIN_OFFICE_CODE:
							loadingOfc.setSalesMainOfficeCode(currentCell.getStringCellValue());
							break;
						case POL_NODE_CODE:
							loadingOfc.setPolNodeCode(currentCell.getStringCellValue());
							break;
						case POD_NODE_CODE:
							loadingOfc.setPodNodeCode(currentCell.getStringCellValue());
							break;
						case OCN_IPC:
							loadingOfc.setOcnIpc(currentCell.getStringCellValue());
							break;
						case BKG_TEU_ALL:
							loadingOfc.setBkgTeuAll(currentCell.getNumericCellValue());
							break;
						case BKG_WGT_ALL:
							loadingOfc.setBkgWgtAll(currentCell.getNumericCellValue());
							break;
						case BKG_TEU_FIRM:
							loadingOfc.setBkgTeuFirm(currentCell.getNumericCellValue());
							break;
						case BKG_WGT_FIRM:
							loadingOfc.setBkgWgtFirm(currentCell.getNumericCellValue());
							break;
					}
				}
				lstData.add(loadingOfc);
			}
			
			// Group by theo sales main office từ lstData vừa tạo
			result =  groupBy(lstData, LoadingOFC2::getSalesMainOfficeCode);

		}
		catch(Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	// Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
 
        return workbook;
    }
    
    private static <E, K> Map<K, List<E>> groupBy(List<E> list, Function<E, K> keyFunction) {
        return Optional.ofNullable(list)
                .orElseGet(ArrayList::new)
                .stream()
                .collect(Collectors.groupingBy(keyFunction));
    }
}
