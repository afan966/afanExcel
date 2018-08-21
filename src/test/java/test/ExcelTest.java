package test;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.excel.ExcelDataIterator;
import com.excel.ExcelRow;
import com.excel.reader.util.ReaderUtil;
import com.excel.writer.util.WriterUtil;

public class ExcelTest {
	
    @Test
    public void reader(){
    	String file = "D:\\Datas\\201808\\15151517\\custom_4275.xlsx";
		ExcelDataIterator iterator = new ReaderUtil(file,3).iterator();
		for (ExcelRow row : iterator) {
			if(row!=null){
				System.out.println(row);
			}
		}
    }
    
    @Test
    public void writer(){
    	String file = "D:\\1.xls";
		WriterUtil writer = new WriterUtil(file, "第一,再来一个,又来一个".split(","), "姓名,昵称,手机号,身高,年龄".split(","), new Class<?>[]{String.class,String.class,String.class,double.class,int.class});
		List<String[]> dataset = new ArrayList<String[]>();
		dataset.add("陈,afan,15986818531,11.1,11".split(","));
		dataset.add("阿,along,15986815525,22.2,22".split(","));
		dataset.add("涛,lt,15988858531,33.3,33".split(","));
		writer.append(dataset);
		writer.append(dataset);

		dataset.add("涛2,lt,15988858531,33.3,33".split(","));
		writer.append(dataset);
		dataset.add("涛3,lt,15988858531,33.333,33".split(","));
		dataset.add("涛3,lt,15988858531,333.333,323".split(","));
		writer.append(1,dataset);
		dataset.clear();
		dataset.add("涛44,lt,15988858531,33.333,33".split(","));
		dataset.add("44,lt,15988858531,333.3333,323".split(","));
		writer.append("又来一个",dataset);
		writer.append("又来一个",dataset);
		writer.append("又来一个",dataset);
		System.out.println("writer success...");
    }
}
