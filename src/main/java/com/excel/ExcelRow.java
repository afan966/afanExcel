package com.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelRow implements Row{
	
	private List<Cell> cells = new ArrayList<Cell>();
	private int rowNum;
	private boolean empty = true;
	
	public boolean isEmpty() {
		return empty;
	}

	public void full() {
		this.empty = false;
	}

	public ExcelRow(int rowNum) {
		this.rowNum = rowNum;
		this.empty = true;
	}
	
	public void addCell(Cell cell){
		cells.add(cell);
	}
	

	@Override
	public Iterator<Cell> iterator() {
		return cells.iterator();
	}

	@Override
	public Iterator<Cell> cellIterator() {
		return cells.iterator();
	}

	@Override
	public Cell createCell(int arg0) {
		return null;
	}

	@Override
	public Cell createCell(int arg0, int arg1) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Cell getCell(int arg0) {
		if(arg0<cells.size()){
			return cells.get(arg0);
		}
		return null;
	}

	@Override
	public Cell getCell(int arg0, MissingCellPolicy arg1) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public short getFirstCellNum() {
		return 0;
	}

	@Override
	public short getHeight() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public float getHeightInPoints() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public short getLastCellNum() {
		return (short)(cells.size()+1);
	}

	@Override
	public int getOutlineLevel() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getPhysicalNumberOfCells() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getRowNum() {
		return rowNum;
	}

	@Override
	public CellStyle getRowStyle() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Sheet getSheet() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public boolean getZeroHeight() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public boolean isFormatted() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void removeCell(Cell arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setHeight(short arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setHeightInPoints(float arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setRowNum(int arg0) {
		this.rowNum = arg0;
	}

	@Override
	public void setRowStyle(CellStyle arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setZeroHeight(boolean arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		for (Cell cell : cells) {
			sb.append(cell.getStringCellValue());
			sb.append(",");
		}
		return sb.toString();
	}
}
