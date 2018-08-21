package com.excel.reader;

import com.excel.ExcelRow;
import com.excel.reader.handler.ExcelDataQueue;

/**
 * Excel��ȡ��֧�ֳ����ļ���ȡ��
 * ֧�ֱ�׼��ʽ��xls,xlsx ֱ�Ӹ��ļ���׺�Ĳ�֧��!!!
 * ����һ����ȫ�����ڴ���
 * ���ݷֶζ�ȡ�������ڴ�����
 * ֻ���ļ�������ɲ�֪��������
 * 
 * @author afan
 * 
 */
public class ExcelDataReader {

	private static final String XLS = ".xls";
	private static final String XLSX = ".xlsx";
	private static final int queueSize = 1024;
	private static final int maxBlankRow = 3;
	ExcelDataQueue.ExlQueue queue;
	private int maxRowCount = 0;
	String name;
	int blankRow = 0;
	int count = 0;
	ExcelXlsReader excelXls = null;
	ExcelXlsxReader excelXlsx = null;
	

	public ExcelDataReader(ExcelDataQueue.ExlQueue queue) throws Exception {
		this.queue = queue;
		this.name = queue.getName();
		this.maxRowCount = queue.getTotal();
		if (name.toLowerCase().endsWith(XLS)) {
			this.excelXls = new ExcelXlsReader();
		} else if (name.toLowerCase().endsWith(XLSX)) {
			this.excelXlsx = new ExcelXlsxReader();
		} else {
			throw new Exception();
		}

	}

	public void add(ExcelRow row) {
		synchronized (queue) {
			try {
				if ((blankRow >= maxBlankRow) || (maxRowCount > 0 && count >= maxRowCount)) {
					if (name.toLowerCase().endsWith(XLS)) {
						excelXls.markEnd();
					} else if (name.toLowerCase().endsWith(XLSX)) {
						excelXlsx.markEnd();
					}
					queue.markEnd();
					return;
				}

				if (row.isEmpty()) {
					blankRow++;
					return;
				}
				if (queue.size() >= queueSize) {
					try {
						queue.wait();
					} catch (Exception e) {
					}
				}

				try {
					queue.add(row);
					count++;
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			} finally {
				queue.notifyAll();
			}
		}
	}

	public void work() {
		try {
			int count = 0;
			synchronized (queue) {
				try {
					if (name.toLowerCase().endsWith(XLS)) {
						count = excelXls.process(name);
					} else if (name.toLowerCase().endsWith(XLSX)) {
						count = excelXlsx.process(name);
					}
					System.out.println(name + "<<<===" + count);
				} catch (Exception e) {
					e.printStackTrace();
				}
				queue.markEnd();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
