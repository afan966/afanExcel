package com.excel.reader;

import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
import com.excel.ExcelCell;
import com.excel.ExcelRow;
import com.excel.reader.handler.ExcelDataCollector;

/**
 * SAXΩ‚ŒˆXLSX
 * @author afan
 *
 */
public class ExcelXlsxReader extends DefaultHandler {
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}
	
	private SharedStringsTable sst;
	private String lastIndex;
	private String filePath = "";
	@SuppressWarnings("unused")
	private int sheetIndex = 0;
	@SuppressWarnings("unused")
	private String sheetName = "";
	private int totalRows = 0;
	private int curRow = 1;
	private ExcelRow row = new ExcelRow(curRow);
	private int curCol = 0;
	private boolean isTElement;
	private String exceptionMessage;
	private CellDataType nextDataType = CellDataType.SSTINDEX;
	private final DataFormatter formatter = new DataFormatter();
	private short formatIndex;
	private String formatString;
	private String preRef = null, ref = null;
	private String maxRef = null;
	InputStream currSheet = null;
	

	private boolean active = true;

	public void markEnd() {
		active = false;
	}
	private StylesTable stylesTable;

	public int process(String filename) throws Exception {
		try {
			filePath = filename;
			OPCPackage pkg = OPCPackage.open(filename);
			XSSFReader xssfReader = new XSSFReader(pkg);
			stylesTable = xssfReader.getStylesTable();
			SharedStringsTable sst = xssfReader.getSharedStringsTable();
			XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
			this.sst = sst;
			parser.setContentHandler(this);
			XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
			if (sheets.hasNext()) {
				try {
					curRow = 1;
					sheetIndex++;
					currSheet = sheets.next();
					sheetName = sheets.getSheetName();
					InputSource sheetSource = new InputSource(currSheet);
					parser.parse(sheetSource);
				} catch (Exception e) {
				} finally {
					currSheet.close();
				}
			}
		} catch (Exception e) {
		}
		return totalRows;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		if ("c".equals(name)) {
			if (preRef == null) {
				preRef = attributes.getValue("r");
			} else {
				preRef = ref;
			}
            ref = attributes.getValue("r");
            int startCellNo = covertRowIdtoCellNo(ref);
            curRow = covertRowIdtoRowNo(ref);
            row.setRowNum(curRow);

            for (int i = 0; i < startCellNo - 1 - curCol; i++) {
                row.addCell(new ExcelCell(row, curRow, "", curCol, HSSFCell.CELL_TYPE_STRING));
            }

            curCol = startCellNo - 1;

			this.setNextDataType(attributes);
		}
		if ("t".equals(name)) {
			isTElement = true;
		} else {
			isTElement = false;
		}
		lastIndex = "";
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		lastIndex += new String(ch, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		if (isTElement) {
			String value = lastIndex.trim();
			row.addCell(new ExcelCell(row, curRow, value, curCol, HSSFCell.CELL_TYPE_STRING));
			curCol++;
			isTElement = false;
			checkRowIsNull(value);
		} else if ("v".equals(name)) {
			String value = this.getDataValue(lastIndex.trim(), "");
//			if (!ref.equals(preRef)) {
//                int len = countNullCell(ref, preRef);
//				for (int i = 0; i < len; i++) {
//                    row.addCell(new ExcelCell(row, curRow, "", curCol, HSSFCell.CELL_TYPE_STRING));
//					curCol++;
//				}
//			}
			row.addCell(new ExcelCell(row, curRow, value, curCol, HSSFCell.CELL_TYPE_STRING));
			curCol++;
			checkRowIsNull(value);
		} else {
			if ("row".equals(name)) {
				if (curRow == 1) {
					maxRef = ref;
				}
				if (maxRef != null) {
					int len = countNullCell(maxRef, ref);
					for (int i = 0; i <= len; i++) {
						row.addCell(new ExcelCell(row, curRow, "", curCol, HSSFCell.CELL_TYPE_STRING));
						curCol++;
					}
				}

				if (active) {
					ExcelDataCollector.get().add(row, filePath);
					totalRows++;
				} else {
					try {
						currSheet.close();
					} catch (IOException e) {
					}
				}

				curCol = 0;
				preRef = null;
				ref = null;
				row = new ExcelRow(curRow);
			}
		}
	}

	public void setNextDataType(Attributes attributes) {
		nextDataType = CellDataType.NUMBER;
		formatIndex = -1;
		formatString = null;
		String cellType = attributes.getValue("t");
		String cellStyleStr = attributes.getValue("s");

		if ("b".equals(cellType)) {
			nextDataType = CellDataType.BOOL;
		} else if ("e".equals(cellType)) {
			nextDataType = CellDataType.ERROR;
		} else if ("inlineStr".equals(cellType)) {
			nextDataType = CellDataType.INLINESTR;
		} else if ("s".equals(cellType)) {
			nextDataType = CellDataType.SSTINDEX;
		} else if ("str".equals(cellType)) {
			nextDataType = CellDataType.FORMULA;
		}

		if (cellStyleStr != null) {
			int styleIndex = Integer.parseInt(cellStyleStr);
			XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
			formatIndex = style.getDataFormat();
			formatString = style.getDataFormatString();
			if (formatString.contains("m/d/yy")) {
				nextDataType = CellDataType.DATE;
				formatString = "yyyy-MM-dd hh:mm:ss";
			}
			if (formatString == null) {
				nextDataType = CellDataType.NULL;
				formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
			}
		}
	}

	public String getDataValue(String value, String thisStr) {
		switch (nextDataType) {
		case BOOL:
			char first = value.charAt(0);
			thisStr = first == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR:
			thisStr = "\"ERROR:" + value.toString() + '"';
			break;
		case FORMULA:
			thisStr = '"' + value.toString() + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
			thisStr = rtsi.toString();
			rtsi = null;
			break;
		case SSTINDEX:
			String sstIndex = value.toString();
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));
				thisStr = rtss.toString();
				rtss = null;
			} catch (NumberFormatException ex) {
				thisStr = value.toString();
			}
			break;
		case NUMBER:
			if (formatString != null) {
				thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
			} else {
				thisStr = value;
			}
			thisStr = thisStr.replace("_", "").trim();
			break;
		case DATE:
			thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
			thisStr = thisStr.replace("T", " ");
			break;
		default:
			thisStr = " ";
			break;
		}
		return thisStr;
	}

	public int countNullCell(String ref, String preRef) {
		String xfd = ref.replaceAll("\\d+", "");
		String xfd_1 = preRef.replaceAll("\\d+", "");
		xfd = fillChar(xfd, 3, '@', true);
		xfd_1 = fillChar(xfd_1, 3, '@', true);
		char[] letter = xfd.toCharArray();
		char[] letter_1 = xfd_1.toCharArray();
		int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
		return res - 1;
	}

	public String fillChar(String str, int len, char let, boolean isPre) {
		int len_1 = str.length();
		if (len_1 < len) {
			if (isPre) {
				for (int i = 0; i < (len - len_1); i++) {
					str = let + str;
				}
			} else {
				for (int i = 0; i < (len - len_1); i++) {
					str = str + let;
				}
			}
		}
		return str;
	}

	public String getExceptionMessage() {
		return exceptionMessage;
	}
	
	public void checkRowIsNull(String value) {
		if (value != null && !"".equals(value)) {
			row.full();
		}
	}
	
	public static int covertRowIdtoCellNo(String rowId){
        int firstDigit = -1;
        for (int c = 0; c < rowId.length(); ++c) {
            if (Character.isDigit(rowId.charAt(c))) {
                firstDigit = c;
                break;
            }
        }
        String newRowId = rowId.substring(0,firstDigit);
        int num = 0;
        int result = 0;
        int length = newRowId.length();
        for(int i = 0; i < length; i++) {
            char ch = newRowId.charAt(length - i - 1);
            num = (int)(ch - 'A' + 1) ;
            num *= Math.pow(26, i);
            result += num;
        }
        return result;
    }
	
	public static int covertRowIdtoRowNo(String rowId){
        int firstDigit = -1;
        for (int c = 0; c < rowId.length(); ++c) {
            if (Character.isDigit(rowId.charAt(c))) {
                firstDigit = c;
                break;
            }
        }
        return Integer.parseInt(rowId.substring(firstDigit));
    }
}