package com.yu.handler;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * @Description Excel处理工具类，包括excel通用导入和通用导出功能
 *
 * @author yu
 * @create 2018/8/20
 * @since 1.0.0
 */
public class PoiExportHandler<T> extends PoiHandler {
  /**
   * 导出excel
   *
   * @param headerStyle 导出表头样式
   * @param dataStyle 导出数据样式
   * @param sheet 导出表格sheet页
   * @param dtoList 要导出的数据
   * @param headers 要导出的表头
   * @throws Exception
   */
  public void exportExcel(
      XSSFCellStyle headerStyle,
      XSSFCellStyle dataStyle,
      Sheet sheet,
      List<T> dtoList,
      String[] headers)
      throws Exception {
    // 设置表头
    setHeaders(headerStyle, sheet, headers);
    // 填充内容
    fillContent(dataStyle, sheet, dtoList);
  }

  /** 设置表头 */
  private void setHeaders(XSSFCellStyle headerStyle, Sheet sheet, String[] headers)
      throws Exception {
    // 生成表头
    Row row = sheet.createRow(0);

    // 填写表头
    for (int i = 0; i < headers.length; i++) {
      Cell cell = row.createCell(i);
      if (headerStyle != null) {
        cell.setCellStyle(headerStyle);
      }
      cell.setCellValue(headers[i]);
    }
  }

  /**
   * 填充数据
   *
   * @param sheet
   * @param dtoList
   */
  private void fillContent(XSSFCellStyle dataStyle, Sheet sheet, List<T> dtoList) throws Exception {
    Iterator<T> recordIter = dtoList.iterator();
    while (recordIter.hasNext()) {
      // 获取最后一行索引值，在该行后添加数据
      int rowIndex = sheet.getPhysicalNumberOfRows();
      Row row = sheet.createRow(rowIndex);

      // 获取实例类所包含的方法名称
      T record = recordIter.next();
      Field[] fields = record.getClass().getDeclaredFields();

      // 顺序写入excel，实体类中的字段和表格标题一一对应
      for (int i = 0; i < fields.length; i++) {
        Cell cell = row.createCell(i);
        if (dataStyle != null) {
          cell.setCellStyle(dataStyle);
        }
        // 获取属性名称，填充到单元格中
        cell.setCellValue(getFieldValue(getMethodName(fields[i]), record));
      }
    }
  }

  /**
   * 获取属性值
   *
   * @param methodName 字段get方法名称
   * @param t 数据vo
   * @return
   */
  private String getFieldValue(String methodName, Object t) {
    Class tCls = t.getClass();
    String fieldValue = "";
    try {
      Method getMethod = tCls.getMethod(methodName, new Class[] {});
      // 通过JavaBean对象拿到该属性的get方法，从而进行操控
      Object val = getMethod.invoke(t, new Object[] {});
      // 操控该对象属性的get方法，从而拿到属性值
      if (val != null) {
        fieldValue = String.valueOf(val);
      }
    } catch (SecurityException
        | InvocationTargetException
        | NoSuchMethodException
        | IllegalAccessException e) {
      e.printStackTrace();
    } catch (IllegalArgumentException e) {
      e.printStackTrace();
    }
    return fieldValue;
  }
}
