package com.excel;

public class ExcelSheet {

	private String title;// sheetName
	private int count;// 行号
	private int sheetNo;// 编号
	private int offset;// 偏移

	private String[] headers;// 表头
	private Class<?>[] valueTypes;// 数据类型

	public ExcelSheet(String title) {
		this.title = title;
	}

	public ExcelSheet(String title, String[] headers) {
		this.title = title;
		this.headers = headers;
		this.incr();
	}

	public boolean hasHeader() {
		return headers != null && headers.length > 0;
	}

	public int getHeaderSize() {
		if (hasHeader()) {
			return headers.length;
		}
		return 0;
	}

	public String getHeader(int idx) {
		if (headers != null && idx >= 0 && headers.length > idx) {
			return headers[idx];
		}
		return null;
	}

	public Class<?> getValueType(int idx) {
		if (valueTypes != null && idx >= 0 && valueTypes.length > idx) {
			return valueTypes[idx];
		}
		return String.class;
	}

	public int incr() {
		return this.count++;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public int getCount() {
		return count;
	}

	public void setCount(int count) {
		this.count = count;
	}

	public int getSheetNo() {
		return sheetNo;
	}

	public void setSheetNo(int sheetNo) {
		this.sheetNo = sheetNo;
	}

	public String[] getHeaders() {
		return headers;
	}

	public void setHeaders(String[] headers) {
		this.headers = headers;
		this.setCount(1);
	}

	public Class<?>[] getValueTypes() {
		return valueTypes;
	}

	public void setValueTypes(Class<?>[] valueTypes) {
		this.valueTypes = valueTypes;
	}

	public int getOffset() {
		return offset;
	}

	public void setOffset(int offset) {
		this.offset = offset;
	}
}
