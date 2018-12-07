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
		dataset.add("涛44,lt,15988858531,33.333,33".split(","));
		dataset.add("44,lt,15988858531,333.3333,323".split(","));
		
		for (int i = 0; i < 20; i++) {
			dataset.add((i+"44,lt,15988858531,333.3333,323").split(","));
		}
//		
//		for (int i = 0; i < 1000; i++) {
//			writer.append(dataset);
//		}
		writer.append("又来一个",dataset);
//		writer.append("又来一个",dataset);
//		writer.append("又来一个",dataset);
		System.out.println("writer success...");
    	} catch (Exception e) {
		}
    }
    
    //@Test
    public void writerSplit(){
    	String file = "D:\\1.xls";
    	try {
//    		WriterUtil writer = new WriterUtil(file, "姓名,昵称,手机号,身高,年龄".split(",")).maxRow(12).subSuffix("");
//        	List<String[]> dataset = new ArrayList<String[]>();
//        	for (int i = 0; i < 10; i++) {
//    			dataset.add((i+"44,lt,15988858531,333.3333,323").split(","));
//    		}
//        	writer.append(dataset);
//        	writer.append(dataset);
    		
    		String header = "取号时间, 运单号, 运单状态, 物流状态, 业务员, 外部编码（客户简称）, 客户名/电子面单账号, 使用者, 目的地, 重量, 运费, 品类, 寄件人姓名, 寄件人电话, 寄件人地址, 寄件人公司, 收件人姓名, 收件人电话, 收件人地址";
    		String header2 = "取号时间2, 运单号2, 运单状态, 物流状态, 业务员, 外部编码（客户简称）, 客户名/电子面单账号, 使用者, 目的地, 重量, 运费, 品类, 寄件人姓名, 寄件人电话, 寄件人地址, 寄件人公司, 收件人姓名, 收件人电话, 收件人地址";
    		String[] headers = header.split(",");
    		String[] headers2 = header2.split(",");
    		Class<?>[] valueTypes = new Class<?>[headers.length];
        	valueTypes[9] = double.class;
        	valueTypes[10] = double.class;
    		WriterUtil writerUtil = new WriterUtil(file, "面单明细", headers, valueTypes);
    		List<String[]> dataset = new ArrayList<String[]>();
    		dataset.add("2018-03-30 18:15:27, 152240492768671, 有效, , , , IT部测试, , 天津, 0.00, 0.00, 服饰, 互范文芳, 18728564512, 山西省阳泉市爱过后人感觉惹她了苦尽甘来看见人头合力科技和福尔恢复和日偶发而听会歌让我体会, , 申通快递, 15875645879, 天津天津市二个人".split(","));
    		dataset.add("2018-03-30 18:15:27, 152240492768671, 有效, , , , IT部测试, , 天津, 0.00, 0.00, 服饰, 互范文芳, 18728564512, 山西省阳泉市爱过后人感觉惹她了苦尽甘来看见人头合力科技和福尔恢复和日偶发而听会歌让我体会, , 申通快递, 15875645879, 天津天津市二个人".split(","));
    		writerUtil.append(dataset);
    		writerUtil.append("面单明细22", headers2, valueTypes, dataset);
    		writerUtil.append("面单明细23", dataset);
    		writerUtil.append("面单明细22", dataset);
		} catch (Exception e) {
			e.printStackTrace();
		}
    	
    }
    	
}
