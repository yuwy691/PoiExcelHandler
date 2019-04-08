package com.yu.handlers;

import java.util.ArrayList;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * @Description 解析类
 *
 * @author yu
 * @date 2019/3/27 16:25
 * @since 1.0.0
 */
public class SaxExcelParse extends DefaultHandler {

  /** 第一行 */
  public static final String FIRST_ROW = "1";
  /** 单元格标签 */
  public static final String XML_C_TAG = "c";
  /** 单元格坐标属性 */
  public static final String XML_R_ATTRIBUTE = "r";
  /** 单元格类型属性 */
  public static final String XML_T_ATTRIBUTE = "t";
  /** Excel行标签 */
  public static final String XML_ROW_TAG = "row";
  /** 单元格类型值，字符串 */
  public static final String XML_S_VALUE = "s";
  /** Field lastContents: 单元格内容 */
  private String lastContents;
  /** 读取Excel数据集 */
  private List<List<String>> rows;
  /** excel对应每一行的值 */
  private List<String> row;
  /** 当前单元格坐标 */
  private String currentKey;
  /** 上一个单元格坐标 */
  private String lastKey;
  /** 是否有表头 */
  private boolean hasHeader = true;
  /** 是否使用sst索引 */
  private boolean sstIndex = false;
  /** 共享字符串 */
  private SharedStringsTable sst;

  /**
   * 构造器
   *
   * @param hasHeader
   * @param sst
   */
  public SaxExcelParse(SharedStringsTable sst, boolean hasHeader) {
    this.hasHeader = hasHeader;
    this.sst = sst;
    this.rows = new ArrayList<>();
    this.row = new ArrayList<>();
  }

  @Override
  public void startElement(String uri, String localName, String name, Attributes attributes)
      throws SAXException {
    if (XML_C_TAG.equals(name)) {
      String cellType = attributes.getValue(XML_T_ATTRIBUTE);
      if (cellType != null && XML_S_VALUE.equals(cellType)) {
        this.sstIndex = true;
      } else {
        this.sstIndex = false;
      }
      this.currentKey = attributes.getValue(XML_R_ATTRIBUTE);
    }
    // 清空内容
    this.lastContents = "";
  }

  /**
   * Title: endElement
   *
   * <p>Description:
   *
   * @param uri uri
   * @param localName localName
   * @param name XML标签名
   * @throws SAXException SAX异常
   * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String, java.lang.String,
   *     java.lang.String)
   */
  @Override
  public void endElement(String uri, String localName, String name) throws SAXException {
    getContentsBySSTIndex();

    if (XML_C_TAG.equals(name)) {
      // 补全空单元格，再处理当前单元格
      if (currentAndLastDiff() > 1) {
        for (int i = 1; i < currentAndLastDiff(); i++) {
          this.row.add("");
        }
      }
      this.row.add(this.lastContents);
      /** 换行设置上一单元格坐标 */
      this.lastKey = this.currentKey;
    } else if (XML_ROW_TAG.equals(name)) {
      if (!ignoreFirstRow()) {
        List<String> item = new ArrayList<>(this.row);
        this.rows.add(item);
      }
      /** 清空存储集 * */
      this.row.clear();
      this.lastKey = "";
    }
  }

  private void getContentsBySSTIndex() {
    if (this.sstIndex) {
      try {
        int idx = Integer.parseInt(lastContents);
        lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      } catch (Exception e) {

      }
    }
  }

  private boolean ignoreFirstRow() {
    String nowRow = this.currentKey.replaceAll(this.currentKey.replaceAll("\\d+", ""), "");
    if (this.hasHeader && FIRST_ROW.equals(nowRow)) {
      return true;
    }
    return false;
  }

  private int currentAndLastDiff() {
    if (StringUtils.isNotBlank(this.lastKey)) {
      return sumStrAscii(this.currentKey) - sumStrAscii(this.lastKey);
    }
    return 0;
  }

  private int sumStrAscii(String str) {
    int sum = 0;
    for (int i = 0; i < str.toCharArray().length; i++) {
      sum += str.toCharArray()[i];
    }
    return sum;
  }

  public List<List<String>> getRows() {
    return this.rows;
  }

  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    this.lastContents += new String(ch, start, length);
  }
}
