package com.excel;

import java.util.ArrayList;
import java.util.List;

public class ExcelSheet {

	public static final int MAX_ROW = 65535;

	private String title;// sheetName
	private int totalCount;// 总行号
	private int sheetNo;// 编号
	private int offset;// 偏移

	private String[] headers;// 表头
	private Class<?>[] valueTypes;// 数据类型

	private List<ExcelSheet> subSheetList;// 超过MAX_ROW新建一个Sheet
	private int currSubSheetNo = 0;// 当前子Sheet的编号
	private int currSubSheetCount;// 当前子Sheet的行号

	public ExcelSheet(String title) {
		this.title = title;
	}

	public ExcelSheet(String title, String[] headers) {
		this.title = title;
		this.headers = headers;
		this.incr();
	}

	public ExcelSheet(ExcelSheet sheet) {
		this(sheet, null);
	}
	
	public ExcelSheet(ExcelSheet sheet, String title) {
		if (title != null) {
			this.title = title;
		} else {
			this.title = sheet.getTitle() + "扩展（" + sheet.getCurrSubSheetNo() + "）";
		}

		this.headers = sheet.getHeaders();
		this.valueTypes = sheet.getValueTypes();
		this.offset = sheet.getOffset();
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

	public void addSubSheet(ExcelSheet subSheet) {
		if (subSheetList == null) {
			subSheetList = new ArrayList<ExcelSheet>();
		}
		subSheetList.add(subSheet);
	}

	public int incr() {
		incrTotalCount();
		return incrSubSheetCount();
	}

	public int incrTotalCount() {
		return this.totalCount++;
	}

	public int incrSubSheetCount() {
		return this.currSubSheetCount++;
	}

	public void setCount(int count) {
		this.currSubSheetCount = count;
		this.totalCount = count;
	}

	public void incrCurrSubSheetNo() {
		this.currSubSheetNo++;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public int getCurrSubSheetCount() {
		return currSubSheetCount;
	}

	public void setCurrSubSheetCount(int currSubSheetCount) {
		this.currSubSheetCount = currSubSheetCount;
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

	public int getTotalCount() {
		return totalCount;
	}

	public void setTotalCount(int totalCount) {
		this.totalCount = totalCount;
	}

	public List<ExcelSheet> getSubSheetList() {
		return subSheetList;
	}

	public void setSubSheetList(List<ExcelSheet> subSheetList) {
		this.subSheetList = subSheetList;
	}

	public int getCurrSubSheetNo() {
		return currSubSheetNo;
	}

	public void setCurrSubSheetNo(int currSubSheetNo) {
		this.currSubSheetNo = currSubSheetNo;
	}

}
