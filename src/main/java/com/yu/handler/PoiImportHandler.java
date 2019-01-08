package com.yu.handler;

import com.yu.enums.ExcelTypeEnum;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Description Excel处理工具类，包括excel通用导入和通用导出功能
 *
 * @author yu
 * @create 2018/8/20
 * @since 1.0.0
 */
public class PoiImportHandler extends PoiHandler {

  /**
   * 导入服务
   *
   * @param in excel文件输入流
   * @param obj 转换的对象，需要对应一个导入vo，字段顺序同excel表列顺序
   * @param excelType excel类型
   * @return
   * @throws Exception
   */
  public List<Map<String, Object>> importExcel(InputStream in, Object obj, String excelType)
      throws Exception {

    Workbook wb;

    if (ExcelTypeEnum.XLSX.getType().equalsIgnoreCase(excelType)) {
      wb = WorkbookFactory.create(in);
    } else if (ExcelTypeEnum.XLS.getType().equals(excelType)) {
      wb = new XSSFWorkbook(in);
    } else {
      throw new Exception("无法解析excel");
    }

    // 获取表格
    Sheet sheet = wb.getSheetAt(0);

    // 返回结果
    List<Map<String, Object>> result = new ArrayList<>();

    // 遍历行 从下标第一行开始（去除标题）
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);

      if (row != null) {
        // 装载obj
        result.add(dataObj(obj, row));
      }
    }
    return result;
  }
}
