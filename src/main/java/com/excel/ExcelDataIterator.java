package com.excel;

import java.util.Iterator;
import com.excel.reader.handler.ExcelDataQueue;

/**
 * Excel数据迭代器,数据直接从这里循环出来
 * 
 * @author afan
 * 
 */
public class ExcelDataIterator implements Iterable<ExcelRow>, Iterator<ExcelRow> {

	private ExcelDataQueue.ExlQueue queue;
	String name;

	public ExcelDataIterator(ExcelDataQueue.ExlQueue queue) {
		this.queue = queue;
		this.name = queue.getName();
	}

	@Override
	public boolean hasNext() {
		synchronized (queue) {
			return queue.getStatus();
		}
	}

	@Override
	public ExcelRow next() {
		ExcelRow row = null;
		synchronized (queue) {
			try {
				row = queue.get();
			} catch (InterruptedException e) {
				e.printStackTrace();
			} finally {
				queue.notifyAll();
			}
		}
		return row;
	}

	@Override
	public void remove() {
	}

	@Override
	public Iterator<ExcelRow> iterator() {
		return this;
	}

}