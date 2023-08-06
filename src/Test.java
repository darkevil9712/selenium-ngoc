import java.util.HashMap;
import java.util.List;
import java.util.Map;

import model.LoadingOFC2;
import util.ExcelReader;

public class Test {

	public static void main(String[] args) {
		Map<String, List<LoadingOFC2>> result = ExcelReader.readDataFromSheet("Long Haul (at Loading Office 2)");
		System.out.print(result);
	}

}
