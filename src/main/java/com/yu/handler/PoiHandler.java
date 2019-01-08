package com.yu.handler;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * @Description Excel处理基类
 *
 * @author yu
 * @create 2019/1/8
 * @since 1.0.0
 */
public class PoiHandler {
  /**
   * 将excel每一行的数据转换为java对象
   *
   * @param obj
   * @param row
   * @return
   * @throws Exception
   */
  Map<String, Object> dataObj(Object obj, Row row) throws Exception {
    Class<?> rowClazz = obj.getClass();
    Field[] fields = FieldUtils.getAllFields(rowClazz);
    if (fields == null || fields.length < 1) {
      return null;
    }

    // 容器
    Map<String, Object> map = new HashMap<>(50);

    // 注意excel表格字段顺序要和obj字段顺序对齐
    for (int j = 0; j < fields.length; j++) {
      map.put(fields[j].getName(), getVal(row.getCell(j)));
    }
    return map;
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

  /**
   * 获取字段get方法名称
   *
   * @param field
   * @return
   */
  String getMethodName(Field field) {
    String first = field.getName().substring(0, 1);
    String other = field.getName().substring(1);

    return "get" + first.toUpperCase() + other;
  }

  /**
   * 获取excel
   *
   * @return
   * @throws Exception
   */
  public SXSSFWorkbook getWorkbook() throws Exception {
    SXSSFWorkbook wb = new SXSSFWorkbook();
    return wb;
  }
}
