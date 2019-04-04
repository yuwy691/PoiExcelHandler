package com.yu.handlers.impl;

import com.yu.handlers.ExcelBaseHandler;
import com.yu.handlers.ExcelHandler;
import com.yu.handlers.SaxExcelParse;
import com.yu.utils.ReflectUtil;
import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * @Description 导出实现类
 *
 * @author yu
 * @date 2019/3/26 20:43
 * @since 1.0.0
 */
public class ExcelHandlerImpl<T> extends ExcelBaseHandler implements ExcelHandler<T> {

  public static final int MAX_XLS_LENGTH = 65536;

  @Override
  public Workbook exportXLSX(String sheetName, boolean needHeader, boolean freeze, List<T> dataList)
      throws Exception {
    Workbook workbook = new SXSSFWorkbook(1000);
    Sheet sheet = workbook.createSheet(sheetName);
    if (needHeader) {
      fillSheet(sheet, freeze, dataList);
    } else {
      fillSheet(sheet, dataList);
    }
    return workbook;
  }

  @Override
  public Workbook exportXLS(String sheetName, boolean needHeader, boolean freeze, List<T> dataList)
      throws Exception {
    if (dataList.size() > MAX_XLS_LENGTH) {
      throw new Exception("xls导出不能超过65536条");
    }

    Workbook workbook = new HSSFWorkbook();
    Sheet sheet = workbook.createSheet(sheetName);
    if (needHeader) {
      fillSheet(sheet, freeze, dataList);
    } else {
      fillSheet(sheet, dataList);
    }
    return workbook;
  }

  @Override
  public List<T> userModelImport(InputStream in, Class<?> clazz) throws Exception {
    Workbook wb = WorkbookFactory.create(in);
    Sheet sheet = wb.getSheetAt(0);

    List<T> list = new ArrayList<>();
    // 遍历行 从下标第一行开始（去除标题）
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);

      if (row != null) {
        // 装载obj
        list.add((T) toBean(clazz, row));
      }
    }

    return list;
  }

  @Override
  public List<T> saxImport(File file, String sheetName, Class<?> clazz, boolean hasHeader)
      throws Exception {
    // 文件地址
    OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ);
    XSSFReader r = new XSSFReader(pkg);
    // 解析的sheet名称
    XSSFReader.SheetIterator sheetsData = (XSSFReader.SheetIterator) r.getSheetsData();
    XMLReader parser = XMLReaderFactory.createXMLReader();
    SaxExcelParse handler = new SaxExcelParse(r.getSharedStringsTable(), hasHeader);
    parser.setContentHandler(handler);
    // 遍历---获取指定的sheet名称
    List<T> list = new ArrayList<>();
    while (sheetsData.hasNext()) {
      InputStream in = sheetsData.next();
      if (sheetName.equals(sheetsData.getSheetName())) {
        InputSource inputSource = new InputSource(in);
        parser.parse(inputSource);
        // 返回所有的封装结果
        List<List<String>> rows = handler.getRows();
        for (List<String> row : rows) {
          // 装载obj
          list.add((T) toBean(clazz, row));
        }
        in.close();
      }
    }
    if (list.isEmpty()) {
      throw new Exception("解析失败，没有找到相应的sheet表！");
    }
    return list;
  }

  @Override
  public void fillSheet(Sheet sheet, List<T> dataList) throws Exception {
    Iterator<T> recordIterator = dataList.iterator();
    while (recordIterator.hasNext()) {
      T record = recordIterator.next();
      Class clazz = record.getClass();
      Field[] fields = clazz.getDeclaredFields();

      // 获取最后一行索引值，在该行后添加数据
      int rowIndex = sheet.getPhysicalNumberOfRows();
      Row row = sheet.createRow(rowIndex);

      // 顺序写入excel，实体类中的字段和表格标题一一对应
      for (int i = 0; i < fields.length; i++) {
        Cell cell = row.createCell(i);

        // 获取属性名称，填充到单元格中
        cell.setCellValue(ReflectUtil.getFieldValue(ReflectUtil.getMethodName(fields[i]), record));
      }
    }
  }

  @Override
  public void fillSheet(Sheet sheet, boolean freeze, List<T> dataList) throws Exception {
    createHeader(null, sheet, getClazz(dataList), freeze);

    fillSheet(sheet, dataList);
  }

  private Class getClazz(List<T> dataList) throws Exception {
    Iterator<T> recordIterator = dataList.iterator();
    T record = recordIterator.next();
    return record.getClass();
  }
}
