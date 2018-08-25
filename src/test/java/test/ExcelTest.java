package test;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.excel.ExcelDataIterator;
import com.excel.ExcelRow;
import com.excel.reader.util.ReaderUtil;
import com.excel.writer.util.WriterUtil;

public class ExcelTest {
	
   // @Test
    public void reader(){
    	String file = "D:\\Datas\\201808\\15151517\\custom_4275.xlsx";
    	file = "D:\\1.xls";
		ExcelDataIterator iterator = new ReaderUtil(file,1024,2).dataIterator();
		
		for (ExcelRow row : iterator) {
			if(row!=null){
				System.out.println(row);
			}
		}
		
		System.out.println(iterator.getHeadRow());
		System.out.println("=============>>>>");
    }
    
    //@Test
    public void writer(){
    	String file = "D:\\1.xls";
    	try {
		WriterUtil writer = new WriterUtil(file, "��һ,����һ��,����һ��".split(","), "����,�ǳ�,�ֻ���,���,����".split(","), new Class<?>[]{String.class,String.class,String.class,double.class,int.class});
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
		dataset.add("��44,lt,15988858531,33.333,33".split(","));
		dataset.add("44,lt,15988858531,333.3333,323".split(","));
		
		for (int i = 0; i < 20; i++) {
			dataset.add((i+"44,lt,15988858531,333.3333,323").split(","));
		}
//		
//		for (int i = 0; i < 1000; i++) {
//			writer.append(dataset);
//		}
		writer.append("����һ��",dataset);
//		writer.append("����һ��",dataset);
//		writer.append("����һ��",dataset);
		System.out.println("writer success...");
    	} catch (Exception e) {
		}
    }
    
    @Test
    public void writerSplit(){
    	String file = "D:\\1.xls";
    	try {
    		WriterUtil writer = new WriterUtil(file, "����,�ǳ�,�ֻ���,���,����".split(",")).maxRow(12).subSuffix("");
        	List<String[]> dataset = new ArrayList<String[]>();
        	for (int i = 0; i < 10; i++) {
    			dataset.add((i+"44,lt,15988858531,333.3333,323").split(","));
    		}
        	writer.append(dataset);
        	writer.append(dataset);
		} catch (Exception e) {
			e.printStackTrace();
		}
    	
    }
    	
}
