package com.excel;

import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelCell implements Cell {
	
	private Row row;
	private int rowIndex;
	
	private String value;
	private int type;
	private int index;
	
	public ExcelCell(Row row, int rowIndex, String value, int index, int type) {
		this.row = row;
		this.rowIndex = rowIndex;
		this.value = value;
		this.type = type;
		this.index = index;
	}

	@Override
	public CellRangeAddress getArrayFormulaRange() {
		return null;
	}

	@Override
	public boolean getBooleanCellValue() {
		return false;
	}

	@Override
	public int getCachedFormulaResultType() {
		return 0;
	}

	@Override
	public Comment getCellComment() {
		return null;
	}

	@Override
	public String getCellFormula() {
		return null;
	}

	@Override
	public CellStyle getCellStyle() {
		return null;
	}

	@Override
	public int getCellType() {
		return type;
	}

	@Override
	public int getColumnIndex() {
		return index;
	}

	@Override
	public Date getDateCellValue() {
		return null;
	}

	@Override
	public byte getErrorCellValue() {
		return 0;
	}

	@Override
	public Hyperlink getHyperlink() {
		return null;
	}

	@Override
	public double getNumericCellValue() {
		try {
			return Double.parseDouble(value);
		} catch (Exception e) {
		}
		return 0;
	}

	@Override
	public RichTextString getRichStringCellValue() {
		return null;
	}

	@Override
	public Row getRow() {
		return row;
	}

	@Override
	public int getRowIndex() {
		return rowIndex;
	}

	@Override
	public Sheet getSheet() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public String getStringCellValue() {
		return value;
	}

	@Override
	public boolean isPartOfArrayFormulaGroup() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void removeCellComment() {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void removeHyperlink() {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setAsActiveCell() {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellComment(Comment arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellErrorValue(byte arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellFormula(String arg0) throws FormulaParseException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellStyle(CellStyle arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellType(int arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(double arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(Date arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(Calendar arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(RichTextString arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(String arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCellValue(boolean arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setHyperlink(Hyperlink arg0) {
		// TODO Auto-generated method stub
		
	}

}
