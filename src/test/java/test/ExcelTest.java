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
		WriterUtil writer = new WriterUtil(file, "��һ,����һ��,����һ��".split(","), "����,�ǳ�,�ֻ���,����,����".split(","), new Class<?>[]{String.class,String.class,String.class,double.class,int.class});
		List<String[]> dataset = new ArrayList<String[]>();
		dataset.add("��,afan,15986818531,11.1,11".split(","));
		dataset.add("��,along,15986815525,22.2,22".split(","));
		dataset.add("��,lt,15988858531,33.3,33".split(","));
		writer.append(dataset);
		writer.append(dataset);

		dataset.add("��2,lt,15988858531,33.3,33".split(","));
		writer.append(dataset);
		dataset.add("��3,lt,15988858531,33.333,33".split(","));
		dataset.add("��3,lt,15988858531,333.333,323".split(","));
		writer.append(1,dataset);
		dataset.clear();
		dataset.add("��44,lt,15988858531,33.333,33".split(","));
		dataset.add("44,lt,15988858531,333.3333,323".split(","));
		writer.append("����һ��",dataset);
		writer.append("����һ��",dataset);
		writer.append("����һ��",dataset);
		System.out.println("writer success...");
    }
}