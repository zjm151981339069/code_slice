package com.zjm.springBoot.codeSlice.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.util.StringUtils;

/**
 * 导出 Excel操作
 * 
 */
public class ExportExcel<T> {

	/***
	 * 单例模式操作Excel导出
	 */
	@SuppressWarnings("rawtypes")
	private static ExportExcel _instance = null;

	private ExportExcel() {
	}

	@SuppressWarnings("rawtypes")
	public static ExportExcel getInstance() {
		if (_instance == null) {
			_instance = new ExportExcel();
		}
		return _instance;
	}

	/***
	 * 根据模版导出Excel文件
	 * 
	 * @param templetName
	 *            模版名称
	 * @param datas
	 *            待导出的数据集
	 * @param fields
	 *            依次从左到右单元格列对应的 datas T 对象的属性名
	 * @param outputStream
	 *            用于网页输出的输出流
	 * @return
	 */
	@SuppressWarnings("resource")
	public Object[] export(String templetName, List<T> datas, String[] fields, String fileName) {
		try {
			// 读取模板
			HSSFWorkbook workbook = null;
			// 工作表
			HSSFSheet sheet = null;
			// 模板输入流
			InputStream is = null;
			HSSFCellStyle style = null;
			Short rowHeight = 0;
			try {
				String fileBaseDir = templetName + ".xls";
				is = getFileInputStream(fileBaseDir);
				try {
					workbook = new HSSFWorkbook(new POIFSFileSystem(is));
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					return this.getError("获取【" + templetName + "】模版失败，请确认模版是否存在！");
				}
			} catch (Exception e) {
				e.printStackTrace();
				return this.getError("系统发生未知的异常，请重试！");
			}
			if (workbook != null) {
				sheet = workbook.getSheetAt(0);
				if (sheet != null) {
					// excel表格待赋值的列数量
					int sheetColumnCount = -1;
					// excel表格操作数据的起始行index
					int sheetRowIndex = -1;
					// 获取sheet表头，判断列的长度
					for (int rowIndex = 0; rowIndex < 10000; rowIndex++) {
						String cellValue = null;
						if (rowIndex == 0) {
							for (int colIndex = 0; colIndex < 10000; colIndex++) {
								HSSFCell cell = sheet.getRow(rowIndex).getCell(colIndex);
								try {
									cellValue = cell.getStringCellValue();
								} catch (Exception ex) {
									// 说明该单元格已没有数据（即模版的有效列结束）
									sheetColumnCount = colIndex;
									break;
								}
								if (cellValue == null || cellValue.trim().equals("")) {
									// 说明该单元格已没有数据（即模版的有效列结束）
									sheetColumnCount = colIndex;
									break;
								}
							}
						}
						if (sheetColumnCount < 0) {
							// 没有表头或表头获取失败
							return this.getError("导出失败，获取【" + templetName + "】模版表头失败！");
						}
						// 获取有效的数据行
						HSSFCell cell = sheet.getRow(rowIndex).getCell(0);
						try {
							cellValue = cell.getStringCellValue();
						} catch (Exception ex) {
							// 说明该单元格已没有数据（即模版的有效表头行结束）
							sheetRowIndex = rowIndex - 1;
							rowHeight = sheet.getRow(rowIndex).getHeight();
							style = sheet.getRow(rowIndex).getCell(1).getCellStyle();
							break;
						}
						if (cellValue == null || cellValue.trim().equals("")) {
							// 说明该单元格已没有数据（即模版的有效列结束）
							sheetRowIndex = rowIndex - 1;
							rowHeight = sheet.getRow(rowIndex).getHeight();
							style = sheet.getRow(rowIndex).getCell(1).getCellStyle();
							break;
						}
					}
					// 给数据行赋值
					if (datas != null && datas.size() > 0) {
						// 获取实体类的所有属性，返回Field数组
						Field[] tFields = datas.get(0).getClass().getDeclaredFields();
						if (fields == null) {
							fields = new String[sheetColumnCount];
							// 遍历所有属性
							for (int index = 0; index < sheetColumnCount; index++) {
								fields[index] = tFields[index].getName();
							}
						}
						for (int rowIndex = 0; rowIndex < datas.size(); rowIndex++) {
							// 获取第 sheetRowIndex 行 ，colIndex列 单元格对象
							sheetRowIndex++;
							HSSFRow row = sheet.createRow(sheetRowIndex);
							row.setHeight(rowHeight);
							T model = datas.get(rowIndex);
							// 获取在对象中对应的单元格的值
							for (int colIndex = 0; colIndex < sheetColumnCount; colIndex++) {
								HSSFCell cell = row.createCell(colIndex);
								cell.setCellStyle(style);
								String fieldName = fields[colIndex];
								if (fieldName.equals("$序号")) {
									// 当前列为序号
									cell.setCellValue(rowIndex + 1);
									continue;
								}
								Field field = null;
								// 获取属性的类型
								for (int filedIndex = 0; filedIndex < tFields.length; filedIndex++) {
									Field fi = tFields[filedIndex];
									if (fi.getName().equals(fieldName)) {
										field = fi;
										break;
									}
								}
								if (field == null) {
									return this.getError("导出失败，类不包含字段【" + fieldName + "】！");
								}
								// 将属性的首字符大写，方便构造get，set方法
								fieldName = fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
								String type = field.getGenericType().toString();
								Method m = model.getClass().getMethod("get" + fieldName);
								if (type.equals("class java.lang.Integer")) {
									Integer value = (Integer) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								if (type.equals("class java.lang.Short")) {
									Short value = (Short) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								if (type.equals("class java.lang.Float")) {
									Float value = (Float) m.invoke(model);
									// 赋值
									cell.setCellValue(value.toString());
									continue;
								}
								if (type.equals("class java.lang.Double")) {
									Double value = (Double) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								if (type.equals("class java.lang.Boolean")) {
									Boolean value = (Boolean) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								if (type.equals("class java.util.Date")) {
									Date value = (Date) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								if (type.equals("class java.lang.Number")) {
									Double value = (Double) m.invoke(model);
									// 赋值
									cell.setCellValue(value);
									continue;
								}
								// 调用getter方法获取属性值
								String value = (String) m.invoke(model);
								// 赋值
								cell.setCellValue(value);
							}

						}
					}
					// 输出
					try {
						FileOutputStream outputStreams = new FileOutputStream(new File(fileName));
						workbook.write(outputStreams);
						workbook.close();
						outputStreams.close();
						if (is != null) {
							try {
								is.close();
							} catch (IOException e) {
								e.printStackTrace();
							}
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
					return getSuccess(fileName);
				}
			}
			return this.getError("导出失败，【" + templetName + "】模版错误!");
		} catch (Exception ex) {
			ex.printStackTrace();
			return null;
		}
	}

	/**
	 * 获取文件流
	 * 
	 * @param path
	 * @return is 文件流
	 * @throws FileNotFoundException
	 */
	private FileInputStream getFileInputStream(String path) throws FileNotFoundException {
		File file = new File(path);
		if (!file.exists()) {
			return null;
		}
		FileInputStream is = new FileInputStream(file);
		return is;
	}

	/***
	 * 获取错误的返回数据
	 * 
	 * @param errorMsg
	 * @return
	 */
	private Object[] getError(String errorMsg) {
		return new Object[] { false, errorMsg };
	}

	/***
	 * 获取正确的返回数据
	 * 
	 * @param errorMsg
	 * @return
	 */
	private Object[] getSuccess(Object msg) {
		return new Object[] { true, msg };
	}

	@SuppressWarnings("resource")
	public List<T> importdata(String templetName, List<T> datas, String[] fields, String fileName) {
		int length = fields.length;
		try {
			// 读取模板
			HSSFWorkbook workbook = null;
			// 工作表
			HSSFSheet sheet = null;
			// 模板输入流
			InputStream is = null;
			HSSFCellStyle style = null;
			List<T> result = new ArrayList<T>();
			try {
				String fileBaseDir = templetName + ".xls";
				is = getFileInputStream(fileBaseDir);
				try {
					workbook = new HSSFWorkbook(new POIFSFileSystem(is));
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					return null;
				}
			} catch (Exception e) {
				e.printStackTrace();
				return null;
			}
			if (workbook != null) {
				sheet = workbook.getSheetAt(0);
				if (sheet != null) {
					// 获取实体类的所有属性，返回Field数组
					Field[] tFields = datas.get(0).getClass().getDeclaredFields();
					String[] allFields = new String[tFields.length];

					// 遍历所有属性
					for (int index = 0; index < tFields.length; index++) {
						allFields[index] = tFields[index].getName();
					}

					// 获取sheet表头，判断列的长度
					for (int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {
						// 判断是否还有记录
						HSSFRow row = sheet.getRow(rowIndex);
						
						T newInstance = (T) datas.get(0).getClass().getConstructor().newInstance(null);
						
						// 遍历每一行的记录
						for (int colIndex = 0; colIndex < length; colIndex++) {
							HSSFCell cell = sheet.getRow(rowIndex).getCell(colIndex);
							try {
								
								Field field = null;
								String fieldName = fields[colIndex];
								// 获取属性的类型
								for (int filedIndex = 0; filedIndex < tFields.length; filedIndex++) {
									Field fi = tFields[filedIndex];
									if (fi.getName().equals(fieldName)) {
										field = fi;
										break;
									}
								}
							    String type = field.getType().toString();
							    String value = cell.toString();
							    
							    if(StringUtils.isEmpty(value)) {
							    	continue;
							    }
							    
							    // 将属性的首字符大写，方便构造get，set方法
								fieldName = fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
								if (type.equals("class java.lang.Integer")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Integer.class);
									m.invoke(newInstance,new Integer(value.substring(0, value.indexOf("."))));
									continue;
								}
								if (type.equals("class java.lang.Short")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Short.class);
									m.invoke(newInstance, new Integer(value.substring(0, value.indexOf("."))));
									continue;
								}
								if (type.equals("class java.lang.String")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, String.class);
									m.invoke(newInstance, value);
									continue;
								}
								if (type.equals("class java.lang.Float")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Float.class);
									m.invoke(newInstance, new Float(value));
									continue;
								}
								if (type.equals("class java.lang.Double")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Double.class);
									m.invoke(newInstance, new Double(value));
									continue;
								}
								if (type.equals("class java.lang.Boolean")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Boolean.class);
									m.invoke(newInstance, new Boolean(value));
									continue;
								}
								if (type.equals("class java.util.Date")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Date.class);
									m.invoke(newInstance, 1);
									continue;
								}
								if (type.equals("class java.lang.Number")) {
									Method m = datas.get(0).getClass().getMethod("set" + fieldName, Number.class);
									m.invoke(newInstance, new Double(value));
									continue;
								}
							} catch (Exception ex) {
								// 出现异常说明列值为空
								System.out.println(ex);
							}
						}
						result.add(newInstance);
					}
				}
			}
			return result;
		} catch (Exception ex) {
			ex.printStackTrace();
			return null;
		}
	}

}
