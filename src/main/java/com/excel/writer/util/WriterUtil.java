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
 * excel����д�빤�ߣ����sheet65535��
 * 
 * @author afan
 * 
 */
public class WriterUtil {

	private String file;
	private List<ExcelSheet> sheets;

	/**
	 * ����sheet,�ޱ�ͷ
	 * 
	 * @param file
	 */
	public WriterUtil(String file) {
		this(file, null);
	}

	/**
	 * ����sheet,����ͷÿ�ж����ַ���
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] headers) {
		this(file, new String[] { "Sheet1" }, new int[1], new String[][] { headers }, null);
	}

	/**
	 * ����sheet,����ͷÿ���������Ͳ�ͬ
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] headers, Class<?>[] valueTypes) {
		this(file, new String[] { "Sheet1" }, new int[1], new String[][] { headers }, new Class<?>[][] { valueTypes });
	}

	/**
	 * ����sheet,����ͷÿ�ж����ַ���
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String sheetName, String[] headers) {
		this(file, new String[] { sheetName }, new int[1], new String[][] { headers }, null);
	}

	/**
	 * ����sheet,����ͷÿ���������Ͳ�ͬ
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String sheetName, String[] headers, Class<?>[] valueTypes) {
		this(file, new String[] { sheetName }, new int[1], new String[][] { headers }, new Class<?>[][] { valueTypes });
	}

	/**
	 * ���sheet��ÿ��sheet��ͬ�ı�ͷ��ÿ�ж����ַ���
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 */
	public WriterUtil(String file, String[] sheetName, String[] headers) {
		String[][] headerss = new String[sheetName.length][headers.length];
		for (int i = 0; i < sheetName.length; i++) {
			headerss[i] = headers;
		}
		this.init(file, sheetName, new int[sheetName.length], headerss, null);
	}

	/**
	 * ���sheet��ÿ��sheet��ͬ�ı�ͷ��ÿ���������Ͳ�ͬ
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 * @param valueTypes
	 */
	public WriterUtil(String file, String[] sheetName, String[] headers, Class<?>[] valueTypes) {
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
	 * ���sheet��ÿ��sheet��ͬ�ı�ͷ��ÿ���������Ͳ�ͬ
	 * 
	 * @param file
	 * @param sheetName
	 * @param headers
	 * @param valueTypes
	 */
	public WriterUtil(String file, String[] sheetName, int[] offset, String[][] headerss, Class<?>[][] valueTypess) {
		this.init(file, sheetName, offset, headerss, valueTypess);
	}

	// ��ʼ������
	private void init(String file, String[] sheetName, int[] offset, String[][] headerss, Class<?>[][] valueTypess) {
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

	// ��ʼ��д���ļ�
	private void createSheets() {
		HSSFWorkbook workbook = new HSSFWorkbook();
		for (ExcelSheet sheet : sheets) {
			createSheet(workbook, sheet);
		}

		OutputStream out = null;
		try {
			out = new FileOutputStream(this.file);
			workbook.write(out);
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
	}

	// ���ɹ�����Sheet
	private void createSheet(HSSFWorkbook workbook, ExcelSheet excelSheet) {
		HSSFSheet sheet = workbook.createSheet(excelSheet.getTitle());
		// Ĭ����ʽ
		sheet.setDefaultColumnWidth(15);
		HSSFCellStyle style = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		style.setFont(font);

		if (excelSheet.hasHeader()) {
			HSSFRow row = sheet.createRow(0);
			for (int i = 0; i < excelSheet.getHeaderSize(); i++) {
				HSSFCell cell = row.createCell(i+excelSheet.getOffset());
				cell.setCellStyle(style);
				HSSFRichTextString text = new HSSFRichTextString(excelSheet.getHeader(i));
				cell.setCellValue(text);
			}
		}
	}

	public int append(List<String[]> dataset) {
		return append(0, dataset);
	}

	public int append(String sheetName, List<String[]> dataset) {
		for (ExcelSheet excelSheet : sheets) {
			if (sheetName.equals(excelSheet.getTitle())) {
				return append(excelSheet, dataset);
			}
		}
		return 0;
	}

	public int append(int sheetNo, List<String[]> dataset) {
		ExcelSheet excelSheet = sheets.get(sheetNo);
		return append(excelSheet, dataset);
	}

	private int append(ExcelSheet excelSheet, List<String[]> dataset) {
		FileInputStream fs = null;
		POIFSFileSystem ps = null;
		HSSFWorkbook wb = null;
		try {
			fs = new FileInputStream(this.file);
			ps = new POIFSFileSystem(fs);
			wb = new HSSFWorkbook(ps);
			HSSFSheet sheet = wb.getSheetAt(excelSheet.getSheetNo());
			HSSFFont font = wb.createFont();
			font.setFontName("yahei");
			HSSFRow row = null;
			Iterator<String[]> it = dataset.iterator();

			while (it.hasNext()) {
				row = sheet.createRow(excelSheet.incr());
				String[] t = it.next();
				for (int i = 0; i < t.length; i++) {
					try {
						HSSFCell cell = row.createCell(i);
						Class<?> clazz = excelSheet.getValueType(i);
						if (clazz != String.class) {
							if (clazz == double.class) {
								cell.setCellValue(Double.parseDouble(t[i]));
							} else if (clazz == long.class) {
								cell.setCellValue(Long.parseLong(t[i]));
							} else if (clazz == int.class) {
								cell.setCellValue(Integer.parseInt(t[i]));
							}
						} else {
							HSSFRichTextString str = new HSSFRichTextString(t[i]);
							str.applyFont(font);
							cell.setCellValue(str);
						}
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
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
		return excelSheet.getCount();
	}

}
