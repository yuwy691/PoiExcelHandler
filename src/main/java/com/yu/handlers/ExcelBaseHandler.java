package com.yu.handlers;

import com.yu.annotations.Header;
import com.yu.exceptions.HeaderException;
import java.lang.reflect.Field;
import java.util.List;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * @Description Excel处理基类
 *
 * @author yu
 * @create 2019/1/8
 * @since 1.0.0
 */
public abstract class ExcelBaseHandler<T> {

  private static final int FREEZE_COLUMN_SPLIT = 0;

  private static final int FREEZE_ROW_SPLIT = 1;

  /**
   * 将excel每一行的数据转换为java对象
   *
   * @param rowClazz
   * @param row
   * @return
   * @throws Exception
   */
  public T toBean(Class<T> rowClazz, Row row) throws Exception {
    Field[] fields = FieldUtils.getAllFields(rowClazz);
    if (fields == null || fields.length < 1) {
      return null;
    }

    // 注意excel表格字段顺序要和obj字段顺序对齐
    T t = rowClazz.newInstance();
    for (int j = 0; j < fields.length; j++) {
      fields[j].setAccessible(true);
      fields[j].set(t, getVal(row.getCell(j)));
    }
    return t;
  }

  public T toBean(Class<T> rowClazz, List<String> row) throws Exception {
    Field[] fields = FieldUtils.getAllFields(rowClazz);
    if (fields == null || fields.length < 1) {
      return null;
    }

    // 注意excel表格字段顺序要和obj字段顺序对齐
    T t = rowClazz.newInstance();
    for (int j = 0; j < fields.length; j++) {
      fields[j].setAccessible(true);
      fields[j].set(t, row.get(j));
    }
    return t;
  }

  /**
   * 创建表头
   *
   * @param headerStyle 表头样式
   * @param sheet sheet页
   * @param clazz 表头字段
   * @param freeze 是否冻结表头
   * @throws Exception
   */
  public void createHeader(XSSFCellStyle headerStyle, Sheet sheet, Class<?> clazz, boolean freeze)
      throws Exception {
    // 生成表头
    Row row = sheet.createRow(0);

    // 是否固定表头
    if (freeze) {
      sheet.createFreezePane(FREEZE_COLUMN_SPLIT, FREEZE_ROW_SPLIT);
    }

    // 设置表头内容
    setHeaderValues(headerStyle, clazz, row);
  }

  private void setHeaderValues(XSSFCellStyle headerStyle, Class<?> clazz, Row row)
      throws HeaderException {
    Field[] fields = clazz.getDeclaredFields();
    for (int i = 0; i < fields.length; i++) {
      Header header = fields[i].getAnnotation(Header.class);
      if (header == null) {
        throw new HeaderException("表头不能为空");
      }
      Cell cell = row.createCell(i);
      if (headerStyle != null) {
        cell.setCellStyle(headerStyle);
      }
      cell.setCellValue(header.value());
    }
  }

  /**
   * 处理val
   *
   * @param cell
   * @return
   */
  String getVal(Cell cell) {
    try {
      String result;
      switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
          result = cell.getStringCellValue();
          break;
        case Cell.CELL_TYPE_NUMERIC:
          result = String.valueOf(cell.getNumericCellValue());
          break;
        default:
          result = cell.getStringCellValue();
      }
      return result;
    } catch (Exception e) {
      return "";
    }
  }
}
