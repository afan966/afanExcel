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
    	String file = "C:\\Users\\afan\\Desktop\\test\\";
    	//file = "D:\\1.xls";
    	
    	tt(file+"1.xls",1);
    	tt(file+"1.xls",1);
    	
//    	for (int i = 1; i < 11; i++) {
//    		tt(file+i+".xls", i);
//		}
    	System.out.println("1111");
//		ExcelDataIterator iterator = new ReaderUtil(file,1024,2).dataIterator();
//		
//		for (ExcelRow row : iterator) {
//			if(row!=null){
//				System.out.println(row);
//			}
//		}
//		
//		System.out.println(iterator.getHeadRow());
//		System.out.println("=============>>>>");
    }
	
	public void tt(String file, int ii){
		ReaderUtil reader = new ReaderUtil(file,50000);
		ExcelDataIterator iterator = reader.iterator();
		
		
		if(reader.iterator().hasNext()){
			System.out.println("-->>"+reader.iterator().next().toString());
		}
		
		try {
			int i =0;
			for (ExcelRow row : iterator) {
				if(row!=null){
					System.out.println(ii+"===>>>"+row);
				}
				if(i>10){
					return;
				}
				i++;
			}
		} catch (Exception e) {
			// TODO: handle exception
		} finally {
			reader.clear();
		}
		
		System.out.println(iterator.getHeadRow());
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
    
    //@Test
    public void writerSplit(){
    	String file = "D:\\1.xls";
    	try {
//    		WriterUtil writer = new WriterUtil(file, "����,�ǳ�,�ֻ���,���,����".split(",")).maxRow(12).subSuffix("");
//        	List<String[]> dataset = new ArrayList<String[]>();
//        	for (int i = 0; i < 10; i++) {
//    			dataset.add((i+"44,lt,15988858531,333.3333,323").split(","));
//    		}
//        	writer.append(dataset);
//        	writer.append(dataset);
    		
    		String header = "ȡ��ʱ��, �˵���, �˵�״̬, ����״̬, ҵ��Ա, �ⲿ���루�ͻ���ƣ�, �ͻ���/�����浥�˺�, ʹ����, Ŀ�ĵ�, ����, �˷�, Ʒ��, �ļ�������, �ļ��˵绰, �ļ��˵�ַ, �ļ��˹�˾, �ռ�������, �ռ��˵绰, �ռ��˵�ַ";
    		String header2 = "ȡ��ʱ��2, �˵���2, �˵�״̬, ����״̬, ҵ��Ա, �ⲿ���루�ͻ���ƣ�, �ͻ���/�����浥�˺�, ʹ����, Ŀ�ĵ�, ����, �˷�, Ʒ��, �ļ�������, �ļ��˵绰, �ļ��˵�ַ, �ļ��˹�˾, �ռ�������, �ռ��˵绰, �ռ��˵�ַ";
    		String[] headers = header.split(",");
    		String[] headers2 = header2.split(",");
    		Class<?>[] valueTypes = new Class<?>[headers.length];
        	valueTypes[9] = double.class;
        	valueTypes[10] = double.class;
    		WriterUtil writerUtil = new WriterUtil(file, "�浥��ϸ", headers, valueTypes);
    		List<String[]> dataset = new ArrayList<String[]>();
    		dataset.add("2018-03-30 18:15:27, 152240492768671, ��Ч, , , , IT������, , ���, 0.00, 0.00, ����, �����ķ�, 18728564512, ɽ��ʡ��Ȫ�а������˸о������˿ྡ����������ͷ�����Ƽ��͸����ָ�����ż����������������, , ��ͨ���, 15875645879, �������ж�����".split(","));
    		dataset.add("2018-03-30 18:15:27, 152240492768671, ��Ч, , , , IT������, , ���, 0.00, 0.00, ����, �����ķ�, 18728564512, ɽ��ʡ��Ȫ�а������˸о������˿ྡ����������ͷ�����Ƽ��͸����ָ�����ż����������������, , ��ͨ���, 15875645879, �������ж�����".split(","));
    		writerUtil.append(dataset);
    		writerUtil.append("�浥��ϸ22", headers2, valueTypes, dataset);
    		writerUtil.append("�浥��ϸ23", dataset);
    		writerUtil.append("�浥��ϸ22", dataset);
		} catch (Exception e) {
			e.printStackTrace();
		}
    	
    }
    	
}
