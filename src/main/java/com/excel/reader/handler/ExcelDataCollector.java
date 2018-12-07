package com.excel.reader.handler;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import com.excel.ExcelRow;
import com.excel.reader.ExcelDataReader;

/**
 * 数据收集器,事件回调的数据提交到这里
 * 
 * @author afan
 * 
 */
public class ExcelDataCollector {

	private static ExcelDataCollector sjq = null;
	private ExecutorService pools = null;
	private static final Map<String, ExcelDataReader> proMap = new HashMap<String, ExcelDataReader>();
	public static int POOL_SIZE = 4;

	public synchronized static ExcelDataCollector get() {
		if (sjq == null) {
			sjq = new ExcelDataCollector();
		}
		return sjq;
	}

	private ExcelDataCollector() {
		pools = Executors.newFixedThreadPool(POOL_SIZE);
	}

	public void addProducer(final ExcelDataQueue.ExlQueue queue) throws Exception {
		synchronized (queue) {
			pools.execute(new Runnable() {

				@Override
				public void run() {
					try {
						ExcelDataReader producer = new ExcelDataReader(queue);
						proMap.put(queue.getName(), producer);
						producer.work();
					} catch (Exception e) {
						e.printStackTrace();
					} finally {
						proMap.remove(queue.getName());
					}
				}
			});
		}
	}

	public void add(ExcelRow row, String name) {
		proMap.get(name).add(row);
	}
	
	public void queueClose(String name){
		if(proMap.get(name)!=null){
			proMap.get(name).queueClose();
		}
	}
	
}
