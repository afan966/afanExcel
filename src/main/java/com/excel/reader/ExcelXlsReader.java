package com.excel.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.record.WindowTwoRecord;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.excel.ExcelCell;
import com.excel.ExcelRow;
import com.excel.reader.handler.ExcelDataCollector;

/**
 * XLS¼àÌýÄ£Ê½
 * @author afan
 *
 */
public class ExcelXlsReader implements HSSFListener {

	private int minColums = -1;
	private POIFSFileSystem fs;
	private int totalRows = 0;
	private int lastRowNumber;
	private int lastColumnNumber;
	private boolean outputFormulaValues = true;
	private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
	private HSSFWorkbook stubWorkbook;
	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;
	private final HSSFDataFormatter formatter = new HSSFDataFormatter();
	private String filePath = "";
	private int sheetIndex = 0;
	private BoundSheetRecord[] orderedBSRs;
	@SuppressWarnings({ "rawtypes" })
	private ArrayList boundSheetRecords = new ArrayList();
	private int nextRow;
	private int nextColumn;
	private boolean outputNextStringRecord;
	private int curRow = 1;
	private ExcelRow row = new ExcelRow(curRow);
	private boolean active = true;
	public void markEnd() {
		active = false;
	}

	@SuppressWarnings("unused")
	private String sheetName;

	public int process(String fileName) throws Exception {
		FileInputStream fis = null;
		try {
			filePath = fileName;
			fis = new FileInputStream(fileName);
			fs = new POIFSFileSystem(fis);
			MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
			formatListener = new FormatTrackingHSSFListener(listener);
			HSSFEventFactory factory = new HSSFEventFactory();
			HSSFRequest request = new HSSFRequest();
			if (outputFormulaValues) {
				request.addListenerForAllRecords(formatListener);
			} else {
				workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
				request.addListenerForAllRecords(workbookBuildingListener);
			}
			factory.processWorkbookEvents(request, fs);
		} catch (Exception e) {
		} finally {
			try {
				fis.close();
				fs.close();
				stubWorkbook.close();
			} catch (Exception e2) {
			}
		}
		return totalRows;
	}

	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String thisStr = null;
		String value = null;
        switch (record.getSid()) {
		case BoundSheetRecord.sid:
			boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				if (workbookBuildingListener != null && stubWorkbook == null) {
					stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
				}

				if (orderedBSRs == null) {
					orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
				}
				sheetName = orderedBSRs[sheetIndex].getSheetname();
				sheetIndex++;
			}
			break;
		case SSTRecord.sid:
			sstRecord = (SSTRecord) record;
			break;
		case BlankRecord.sid:
			BlankRecord brec = (BlankRecord) record;
			thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			thisStr = "";
			row.addCell(new ExcelCell(row, curRow, thisStr, thisColumn, HSSFCell.CELL_TYPE_STRING));
			break;
		case BoolErrRecord.sid:
			BoolErrRecord berec = (BoolErrRecord) record;
			thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			thisStr = berec.getBooleanValue() + "";
			row.addCell(new ExcelCell(row, curRow, thisStr, thisColumn, HSSFCell.CELL_TYPE_STRING));
			checkRowIsNull(thisStr);
			break;
		case FormulaRecord.sid:
			FormulaRecord frec = (FormulaRecord) record;
			thisRow = frec.getRow();
			thisColumn = frec.getColumn();
			if (outputFormulaValues) {
				if (Double.isNaN(frec.getValue())) {
					outputNextStringRecord = true;
					nextRow = frec.getRow();
					nextColumn = frec.getColumn();
				} else {
					thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
				}
			} else {
				thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
			}
			row.addCell(new ExcelCell(row, curRow, thisStr, thisColumn, HSSFCell.CELL_TYPE_STRING));
			checkRowIsNull(thisStr);
			break;
		case StringRecord.sid:
			if (outputNextStringRecord) {
				StringRecord srec = (StringRecord) record;
				thisStr = srec.getString();
				thisRow = nextRow;
				thisColumn = nextColumn;
				outputNextStringRecord = false;
			}
			break;
		case LabelRecord.sid:
			LabelRecord lrec = (LabelRecord) record;
			curRow = thisRow = lrec.getRow();
			thisColumn = lrec.getColumn();
			value = lrec.getValue().trim();
			value = value.equals("") ? "" : value;
			row.addCell(new ExcelCell(row, curRow, value, thisColumn, HSSFCell.CELL_TYPE_STRING));
			checkRowIsNull(value);
			break;
		case LabelSSTRecord.sid:
			LabelSSTRecord lsrec = (LabelSSTRecord) record;
			curRow = thisRow = lsrec.getRow();
			thisColumn = lsrec.getColumn();
			if (sstRecord == null) {
				row.addCell(new ExcelCell(row, curRow, "", thisColumn, HSSFCell.CELL_TYPE_STRING));
			} else {
				value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
				value = value.equals("") ? "" : value;
				row.addCell(new ExcelCell(row, curRow, value, thisColumn, HSSFCell.CELL_TYPE_STRING));
				checkRowIsNull(value);
			}
			break;
		case NumberRecord.sid:
			NumberRecord numrec = (NumberRecord) record;
			curRow = thisRow = numrec.getRow();
			thisColumn = numrec.getColumn();
			Double valueDouble = ((NumberRecord) numrec).getValue();
			String formatString = formatListener.getFormatString(numrec);
			if (formatString.contains("m/d/yy")) {
				formatString = "yyyy-MM-dd hh:mm:ss";
			}
			int formatIndex = formatListener.getFormatIndex(numrec);
			value = formatter.formatRawCellContents(valueDouble, formatIndex, formatString).trim();
			value = value.equals("") ? "" : value;
			row.addCell(new ExcelCell(row, curRow, value, thisColumn, HSSFCell.CELL_TYPE_STRING));
			checkRowIsNull(value);
			break;
		default:
			break;
		}
		if (thisRow != -1 && thisRow != lastRowNumber) {
			lastColumnNumber = -1;
		}
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			curRow = thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			row.addCell(new ExcelCell(row, curRow, "", thisColumn, HSSFCell.CELL_TYPE_STRING));
		}
		if (thisRow > -1)
			lastRowNumber = thisRow;
		if (thisColumn > -1)
			lastColumnNumber = thisColumn;
		if (record instanceof LastCellOfRowDummyRecord || record instanceof WindowTwoRecord) {
			if (minColums > 0) {
				if (lastColumnNumber == -1) {
					lastColumnNumber = 0;
				}
			}
			lastColumnNumber = -1;
			if (active && sheetIndex <= 1) {
				ExcelDataCollector.get().add(row, filePath);
				totalRows++;
			} else {
				close();
			}
			curRow++;
			row = new ExcelRow(curRow);
		}
	}
	
	private void close() {
		try {
			stubWorkbook.close();
		} catch (IOException e) {
		}
		try {
			fs.close();
		} catch (IOException e) {
		}
	}

	public void checkRowIsNull(String value) {
		if (value != null && !"".equals(value)) {
			row.full();
		}
	}
}
