package com.yu.model;

import com.yu.annotations.Header;

/**
 * @Description 导出测试实体类
 *
 * @author yu
 * @create 2019/1/8
 * @since 1.0.0
 */
public class ExportVo {
  @Header("序号")
  private String orderNo;

  @Header("名称")
  private String name;

  @Header("描述")
  private String desc;

  @Header("时间")
  private String exportNow;

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public String getDesc() {
    return desc;
  }

  public void setDesc(String desc) {
    this.desc = desc;
  }

  public String getOrderNo() {
    return orderNo;
  }

  public void setOrderNo(String orderNo) {
    this.orderNo = orderNo;
  }

  public String getExportNow() {
    return exportNow;
  }

  public void setExportNow(String exportNow) {
    this.exportNow = exportNow;
  }
}
