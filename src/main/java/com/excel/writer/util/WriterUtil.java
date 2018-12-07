package com.excel.writer.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.excel.ExcelSheet;

/**
 * excel分批写入工具，最大单sheet65535行
 * 
 * @author afan
 * 
 */
public class WriterUtil {

	private int dataMaxRow = ExcelSheet.MAX_ROW - 1;
	private String subSheetName = "_";
	private String file;
	private List<ExcelSheet> sheets;

	/**
	 * 单个sheet,无表头
	 * 
	 * @param file
	 */
	public WriterUtil(String file) throws IOException {
		this(file, null);
	}

	/**
	 * 单个sheet,带表头每列都是字符串
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] headers) throws IOException {
		this(file, new String[] { "Sheet1" }, new int[1], new String[][] { headers }, null);
	}

	/**
	 * 单个sheet,带表头每列数据类型不同
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] headers, Class<?>[] valueTypes) throws IOException {
		this(file, new String[] { "Sheet1" }, new int[1], new String[][] { headers }, new Class<?>[][] { valueTypes });
	}

	/**
	 * 单个sheet,带表头每列都是字符串
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String sheetName, String[] headers) throws IOException {
		this(file, new String[] { sheetName }, new int[1], new String[][] { headers }, null);
	}

	/**
	 * 单个sheet,带表头每列数据类型不同
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String sheetName, String[] headers, Class<?>[] valueTypes) throws IOException {
		this(file, new String[] { sheetName }, new int[1], new String[][] { headers }, new Class<?>[][] { valueTypes });
	}

	/**
	 * 多个sheet，每个sheet相同的表头，每列都是字符串
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] sheetName, String[] headers) throws IOException {
		String[][] headerss = new String[sheetName.length][headers.length];
		for (int i = 0; i < sheetName.length; i++) {
			headerss[i] = headers;
		}
		this.init(file, sheetName, new int[sheetName.length], headerss, null);
	}

	/**
	 * 多个sheet，每个sheet相同的表头，每列数据类型不同
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 * @param valueTypes
	 */
	public WriterUtil(String file, String[] sheetName, String[] headers, Class<?>[] valueTypes) throws IOException {
		String[][] headerss = new String[sheetName.length][headers.length];
		for (int i = 0; i < sheetName.length; i++) {
			headerss[i] = headers;
		}
		Class<?>[][] valueTypess = new Class<?>[sheetName.length][headers.length];
		for (int i = 0; i < sheetName.length; i++) {
			valueTypess[i] = valueTypes;
		}
		this.init(file, sheetName, new int[sheetName.length], headerss, valueTypess);
	}

	/**
	 * 多个sheet，每个sheet不同的表头，每列数据类型不同
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 * @param valueTypes
	 */
	public WriterUtil(String file, String[] sheetName, int[] offset, String[][] headerss, Class<?>[][] valueTypess) throws IOException {
		this.init(file, sheetName, offset, headerss, valueTypess);
	}

	// 初始化数据
	private void init(String file, String[] sheetName, int[] offset, String[][] headerss, Class<?>[][] valueTypess) throws IOException {
		this.file = file;
		new File(this.file).delete();
		File dirFile = new File(this.file).getParentFile();
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		sheets = new ArrayList<ExcelSheet>();
		for (int i = 0; i < sheetName.length; i++) {
			ExcelSheet sheet = new ExcelSheet(sheetName[i]);
			sheet.setSheetNo(i);
			if (offset != null) {
				sheet.setOffset(offset[i]);
			}
			if (headerss != null) {
				sheet.setHeaders(headerss[i]);
			}
			if (valueTypess != null) {
				sheet.setValueTypes(valueTypess[i]);
			}
			sheets.add(sheet);
		}
		if (sheets != null && sheets.size() > 0) {
			createSheets();
		}
	}
	public WriterUtil maxRow(int maxRow) {
		this.dataMaxRow = maxRow;
		return this;
	}

	public WriterUtil subSuffix(String suffix) {
		this.subSheetName = suffix;
		return this;
	}

	// 初始化写入文件
	private void createSheets() throws IOException {
		HSSFWorkbook workbook = new HSSFWorkbook();
		for (ExcelSheet sheet : sheets) {
			createSheet(workbook, sheet);
		}

		OutputStream out = null;
		try {
			out = new FileOutputStream(this.file);
			workbook.write(out);
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	// 生成工作表Sheet
	private void createSheet(HSSFWorkbook workbook, ExcelSheet excelSheet) {
		HSSFSheet sheet = workbook.createSheet(excelSheet.getTitle());
		sheet.setDefaultColumnWidth(15);
		HSSFCellStyle style = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		style.setFont(font);

		if (excelSheet.hasHeader()) {
			HSSFRow row = sheet.createRow(0);
			for (int i = 0; i < excelSheet.getHeaderSize(); i++) {
				HSSFCell cell = row.createCell(i + excelSheet.getOffset());
				cell.setCellStyle(style);
				HSSFRichTextString text = new HSSFRichTextString(excelSheet.getHeader(i));
				cell.setCellValue(text);
			}
		}
	}

	//追加一个新的sheet
	private void createSubSheet(ExcelSheet subSheet) {
		OutputStream out = null;
		FileInputStream fi = null;
		POIFSFileSystem fs = null;
		HSSFWorkbook wb = null;
		try {
			fi = new FileInputStream(this.file);
			fs = new POIFSFileSystem(fi);
			wb = new HSSFWorkbook(fs);

			createSheet(wb, subSheet);
			out = new FileOutputStream(this.file);
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fi != null) {
					fi.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public int append(List<String[]> dataset) {
		return append(0, dataset);
	}
	
	public int append(String sheetName, List<String[]> dataset) {
		return append(sheetName, null, null, dataset);
	}

	public int append(String sheetName, String[] headers, Class<?>[] valueTypes, List<String[]> dataset) {
		for (ExcelSheet excelSheet : sheets) {
			if (sheetName.equals(excelSheet.getTitle())) {
				return append(excelSheet, dataset);
			}
		}
		
		try {
			//沒有sheet就生成新的
			ExcelSheet excelSheet = createExcelSheet(sheetName, headers, valueTypes);
			return append(excelSheet, dataset);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return 0;
	}
	
	private ExcelSheet createExcelSheet(String sheetName, String[] headers, Class<?>[] valueTypes) throws IOException {
		
		ExcelSheet sheet = new ExcelSheet(sheetName);
		sheet.setSheetNo(sheets.size());
		sheet.setOffset(0);
		if (headers != null) {
			sheet.setHeaders(headers);
		}else{
			sheet.setHeaders(sheets.get(0).getHeaders());
		}
		if (valueTypes != null) {
			sheet.setValueTypes(valueTypes);
		}else{
			sheet.setValueTypes(sheets.get(0).getValueTypes());
		}

		createSubSheet(sheet);
		sheets.add(sheet);
		return sheet;
	}

	public int append(int sheetNo, List<String[]> dataset) {
		if (sheetNo >= 0 && sheetNo < sheets.size()) {
			ExcelSheet excelSheet = sheets.get(sheetNo);
			return append(excelSheet, dataset);
		}
		return 0;
	}

	private int append(ExcelSheet excelSheet, List<String[]> dataset) {
		if (excelSheet.getCurrSubSheetCount() + dataset.size() > dataMaxRow) {
			int remainCount = dataMaxRow - excelSheet.getCurrSubSheetCount();
			List<String[]> subDataset = new ArrayList<String[]>();
			int currCount = 0;
			for (String[] data : dataset) {
				subDataset.add(data);
				currCount++;
				if (currCount >= remainCount) {
					appendSheet(excelSheet, subDataset);
					subDataset.clear();
					// 添加扩展sheet
					excelSheet.incrCurrSubSheetNo();
					if (subSheetName.indexOf("%i") > -1) {
						subSheetName = subSheetName.replaceAll("%i", excelSheet.getCurrSubSheetNo() + "");
					} else {
						subSheetName = subSheetName + excelSheet.getCurrSubSheetNo();
					}
					String subSheetTitle = excelSheet.getTitle() + subSheetName;
					ExcelSheet subSheet = new ExcelSheet(excelSheet, subSheetTitle);
					subSheet.setSheetNo(statSheetCount());
					createSubSheet(subSheet);
					excelSheet.setCurrSubSheetCount(subSheet.getTotalCount() + 1);
					excelSheet.addSubSheet(subSheet);
					remainCount = dataMaxRow - excelSheet.getCurrSubSheetCount();
					currCount = 0;
				}
			}
			if (subDataset.size() > 0) {
				append(excelSheet, subDataset);
			}
		} else {
			return appendSheet(excelSheet, dataset);
		}
		return excelSheet.getTotalCount();
	}

	private int appendSheet(ExcelSheet excelSheet, List<String[]> dataset) {
		FileInputStream fs = null;
		POIFSFileSystem ps = null;
		HSSFWorkbook wb = null;
		try {
			fs = new FileInputStream(this.file);
			ps = new POIFSFileSystem(fs);
			wb = new HSSFWorkbook(ps);

			Class<?>[] valueTypes = excelSheet.getValueTypes();
			HSSFSheet sheet = null;
			if (excelSheet.getCurrSubSheetNo() <= 0) {
				sheet = wb.getSheetAt(excelSheet.getSheetNo());
			} else {
				// 获取子Sheet对应的整个Excel编号
				sheet = wb.getSheetAt(excelSheet.getSubSheetList().get(excelSheet.getCurrSubSheetNo() - 1).getSheetNo());
			}

			Iterator<String[]> it = dataset.iterator();
			while (it.hasNext()) {
				String[] data = it.next();
				appendRow(sheet, valueTypes, data, excelSheet.getCurrSubSheetCount());
				excelSheet.incr();
			}

			OutputStream out = null;
			try {
				out = new FileOutputStream(this.file);
				wb.write(out);
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					if (out != null) {
						out.close();
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (wb != null) {
					wb.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			try {
				if (ps != null) {
					ps.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			try {
				if (fs != null) {
					fs.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return excelSheet.getTotalCount();
	}

	private void appendRow(HSSFSheet sheet, Class<?>[] valueTypes, String[] data, int rowNo) {
		HSSFRow row = null;
		row = sheet.createRow(rowNo);
		for (int i = 0; i < data.length; i++) {
			try {
				HSSFCell cell = row.createCell(i);
				Class<?> clazz = String.class;
				if (valueTypes != null && valueTypes.length > i) {
					clazz = valueTypes[i];
				}
				if (clazz != String.class) {
					if (clazz == double.class) {
						cell.setCellValue(Double.parseDouble(data[i]));
					} else if (clazz == long.class) {
						cell.setCellValue(Long.parseLong(data[i]));
					} else if (clazz == int.class) {
						cell.setCellValue(Integer.parseInt(data[i]));
					} else {
						HSSFRichTextString str = new HSSFRichTextString(data[i]);
						cell.setCellValue(str);
					}
				} else {
					HSSFRichTextString str = new HSSFRichTextString(data[i]);
					cell.setCellValue(str);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private int statSheetCount() {
		int count = sheets.size();
		for (ExcelSheet sheet : sheets) {
			if (sheet.getSubSheetList() != null) {
				count += sheet.getSubSheetList().size();
			}
		}
		return count;
	}

	public static <T> List<List<T>> split(List<T> resList, int count) {
		if (resList == null || count < 1)
			return null;
		List<List<T>> ret = new ArrayList<List<T>>();
		int size = resList.size();
		if (size <= count) {
			ret.add(resList);
		} else {
			int pre = size / count;
			int last = size % count;
			for (int i = 0; i < pre; i++) {
				List<T> itemList = new ArrayList<T>();
				for (int j = 0; j < count; j++) {
					itemList.add(resList.get(i * count + j));
				}
				ret.add(itemList);
			}
			if (last > 0) {
				List<T> itemList = new ArrayList<T>();
				for (int i = 0; i < last; i++) {
					itemList.add(resList.get(pre * count + i));
				}
				ret.add(itemList);
			}
		}
		return ret;

	}
}
