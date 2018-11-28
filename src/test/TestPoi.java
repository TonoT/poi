package test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;

public class TestPoi {
	
	@Test
	public void test() throws Exception{
		InputStream inputStream = new FileInputStream("C:/Users/xzg/Desktop/resourse.xls");
		String suffix ="xls";
		int startrow =2;
	ExcelT t1 = new ExcelT();
	List<String[]> result = t1.parseExcel(inputStream, suffix, startrow);
	for (String[] ss : result) {
		System.out.println(Arrays.toString(ss));
	}
	}
	
	

}
