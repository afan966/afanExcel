package com.excel.reader.handler;

import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.LinkedBlockingQueue;
import com.excel.ExcelRow;

/**
 * ���ݴ�Ŷ��У�ÿ���ļ�һ�������Ķ���
 * 
 * @author afan
 * 
 */
public class ExcelDataQueue {

	static ExcelDataQueue queue = null;

	static final ConcurrentHashMap<String, ExlQueue> exlHashMap = new ConcurrentHashMap<String, ExcelDataQueue.ExlQueue>();

	private static final int DEFAULT_QUEUE_SIZE = 1024;// ���г���
	private static final int DEFAULT_MAX_BLANK_ROW = 3;// ���������Ŀ�����

	private ExcelDataQueue() {
	};

	public static synchronized ExcelDataQueue init() {
		if (queue == null) {
			queue = new ExcelDataQueue();
		}
		return queue;
	}

	public ExlQueue getQueue(String name) {
		return getQueue(name, 0, DEFAULT_QUEUE_SIZE, DEFAULT_MAX_BLANK_ROW);
	}

	public ExlQueue getQueue(String name, int total) {
		return getQueue(name, total, DEFAULT_QUEUE_SIZE, DEFAULT_MAX_BLANK_ROW);
	}

	public ExlQueue getQueue(String name, int total, int queueSize) {
		return getQueue(name, total, queueSize, DEFAULT_MAX_BLANK_ROW);
	}

	public ExlQueue getQueue(String name, int total, int queueSize, int maxBlankRow) {
		synchronized (name) {
			ExlQueue exl = exlHashMap.get(name);
			if (exl == null) {
				exl = new ExlQueue(name, total, queueSize, maxBlankRow);
				exlHashMap.put(name, exl);
			}
			return exl;
		}
	}

	public void clear(String name) {
		synchronized (name) {
			exlHashMap.remove(name);
		}
	}

	public static class ExlQueue {
		private String name = null;// �ļ���
		private int total = 0;// �����ȡ��������
		private int queueSize = 1024;// ���г���
		private int maxBlankRow = 3;// ���������Ŀ�����
		public int status = 0;// 0��ʼ��1��ʼ��ȡ��2�ļ���ȡ����
		private LinkedBlockingQueue<ExcelRow> queue = null;

		private ExlQueue(String name, int total, int queueSize, int maxBlankRow) {
			this.name = name;
			this.status = 0;
			this.total = total;
			this.queueSize = queueSize;
			this.maxBlankRow = maxBlankRow;
			this.queue = new LinkedBlockingQueue<ExcelRow>();
		}

		public synchronized void add(ExcelRow row) throws InterruptedException {
			queue.put(row);
			this.status = 1;
		}

		public synchronized ExcelRow get() throws InterruptedException {
			ExcelRow row = queue.poll();
			markEnd();
			return row;
		}

		public synchronized void markEnd() {
			if (size() == 0) {
				this.status = 2;
			}
		}

		public synchronized boolean getStatus() {
			if (status == 0) {
				try {
					this.wait(1000);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			}
			return status < 2;
		}

		public String getName() {
			return name;
		}

		public int getTotal() {
			return total;
		}

		public int getQueueSize() {
			return queueSize;
		}

		public int getMaxBlankRow() {
			return maxBlankRow;
		}

		public int size() {
			return queue.size();
		}
	}
}
