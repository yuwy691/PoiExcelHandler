package com.yu;

import com.yu.handlers.ExcelHandler;
import com.yu.handlers.impl.ExcelHandlerImpl;
import com.yu.model.ExportVo;
import com.yu.model.ImportVo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

/**
 * @Description 单元测试类
 *
 * @author yu
 * @create 2019/1/8
 * @since 1.0.0
 */
public class HandlerTest {

  @Test
  public void testExport() throws Exception {
    // 导出到本地磁盘
    String path = "D:/test.xlsx";

    List<ExportVo> dtoList = new ArrayList<>();
    for (int i = 0; i < 100; i++) {
      ExportVo vo = new ExportVo();
      vo.setOrderNo(String.valueOf(i+1));
      vo.setName("name" + (i + 1));
      vo.setDesc("desc" + (i + 1));
      vo.setExportNow(Calendar.getInstance().toString());
      dtoList.add(vo);
    }
    ExcelHandler handler = new ExcelHandlerImpl();
    Workbook wb = handler.exportXLSX("name", true, false, dtoList);

    File file = new File(path);
    OutputStream os = new FileOutputStream(file);
    wb.write(os);
  }

  @Test
  public void testImport() throws Exception {
    ExcelHandler handler = new ExcelHandlerImpl();

    String path = "D:/test.xlsx";
    File file = new File(path);
    InputStream is = new FileInputStream(file);
    List<ImportVo> list = handler.saxImport(file, "name", ImportVo.class, true);

    for (ImportVo item : list) {
      System.out.println(item.getDesc());
    }
  }
}
