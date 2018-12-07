package com.excel.reader.util;

import java.io.Closeable;
import java.io.IOException;
import com.excel.ExcelDataIterator;
import com.excel.reader.handler.ExcelDataCollector;
import com.excel.reader.handler.ExcelDataQueue;

/**
 * excel分批读取工具,支持超大文件 
 * 使用方法：
 * ExcelDataIterator iterator = ReaderUtil("xxx.xls").iterator(); 
 * for (ExcelRow row : iterator) { 
 * //do something. 
 * }
 * 第一行是表头,不作为数据处理
 * ExcelDataIterator iterator = ReaderUtil("xxx.xls").dataIterator(); 
 * iterator.getHeadRow();//表头
 * 
 * @author afan
 * 
 */
public class ReaderUtil implements Closeable {

	private ExcelDataQueue queue = ExcelDataQueue.init();
	private ExcelDataCollector collector = ExcelDataCollector.get();
	private ExcelDataIterator iterator = null;
	private String file = null;

	public ReaderUtil(String file) {
		this.init(file, queue.getQueue(file));
	}

	public ReaderUtil(String file, int total) {
		this.init(file, queue.getQueue(file, total));
	}

	public ReaderUtil(String file, int total, int queueSize) {
		this.init(file, queue.getQueue(file, total, queueSize));
	}

	public ReaderUtil(String file, int total, int queueSize, int maxBlankRow) {
		this.init(file, queue.getQueue(file, total, queueSize, maxBlankRow));
	}

	private void init(String file, ExcelDataQueue.ExlQueue exlQueue) {
		try {
			this.file = file;
			collector.addProducer(exlQueue);
			iterator = new ExcelDataIterator(exlQueue);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public ExcelDataIterator iterator() {
		return iterator;
	}
	
	public ExcelDataIterator dataIterator() {
		iterator.getHeadRow();
		return iterator;
	}
	
	public void clear() {
		collector.queueClose(file);
		iterator.queueNotify();
		queue.clear(file);
	}

	@Override
	public void close() throws IOException {
		this.clear();
	}

}
