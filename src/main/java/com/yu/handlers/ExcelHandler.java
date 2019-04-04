package com.yu.handlers;

import java.io.File;
import java.io.InputStream;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Description excel处理服务
 *
 * @author yu
 * @date 2019/3/26 17:09
 * @since 1.0.0
 */
public interface ExcelHandler<T> {

  /**
   * 所有数据导出到一个sheet页，导出xlsx格式
   *
   * @param sheetName sheet名称
   * @param needHeader 是否需要表头
   * @param freeze 是否冻结表头
   * @return
   * @throws Exception
   */
  Workbook exportXLSX(String sheetName, boolean needHeader, boolean freeze, List<T> dataList)
      throws Exception;

  /**
   * 所有数据导出到一个sheet页，导出xls格式
   *
   * @param sheetName sheet名称
   * @param needHeader 是否需要表头
   * @param freeze 是否冻结表头
   * @return
   * @throws Exception
   */
  Workbook exportXLS(String sheetName, boolean needHeader, boolean freeze, List<T> dataList)
      throws Exception;

  /**
   * 正常导入
   *
   * @param in 输入表格文件流
   * @param clazz excel表对应的bean,表格字段顺序要和bean字段顺序对齐
   * @return
   * @throws Exception
   */
  List<T> userModelImport(InputStream in, Class<?> clazz) throws Exception;

  /**
   * 事件驱动方式导入,单元格都是文本类型
   *
   * @param file 要导入的文件
   * @param sheetName sheet名
   * @param clazz excel表对应的bean,表格字段顺序要和bean字段顺序对齐,字段类型都是字符串
   * @param hasHeader 导入的文件是否含有表头
   * @return
   * @throws Exception
   */
  List<T> saxImport(File file, String sheetName, Class<?> clazz, boolean hasHeader)
      throws Exception;

  /**
   * 填充sheet页
   *
   * @param sheet sheet
   * @param dataList 填充数据
   * @throws Exception
   */
  void fillSheet(Sheet sheet, List<T> dataList) throws Exception;

  /**
   * 填充sheet页，带表头
   *
   * @param sheet sheet
   * @param freeze 是否冻结表头
   * @param dataList 填充数据
   * @throws Exception
   */
  void fillSheet(Sheet sheet, boolean freeze, List<T> dataList) throws Exception;
}
