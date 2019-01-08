package com.yu;

import com.yu.handler.PoiExportHandler;
import com.yu.handler.PoiHandler;
import com.yu.handler.PoiImportHandler;
import com.yu.model.ExportVo;
import com.yu.model.ImportVo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Sheet;
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
    String[] headers = {"序号", "名称", "描述"};
    List<ExportVo> dtoList = new ArrayList<>();
    for (int i = 0; i < 10; i++) {
      ExportVo vo = new ExportVo();
      vo.setOrderNo(String.valueOf(i + 1));
      vo.setName("name" + i);
      vo.setDesc("desc" + i);

      dtoList.add(vo);
    }
    PoiHandler baseHandler = new PoiHandler();
    PoiExportHandler handler = new PoiExportHandler();

    Workbook wb = baseHandler.getWorkbook();
    Sheet sheet = wb.createSheet("test");
    handler.exportExcel(null, null, sheet, dtoList, headers);

    File file = new File(path);
    OutputStream os = new FileOutputStream(file);

    wb.write(os);
  }

  @Test
  public void testImport() throws Exception {
    PoiImportHandler handler = new PoiImportHandler();

    String path = "D:/test.xlsx";
    File file = new File(path);
    InputStream is = new FileInputStream(file);
    List<Map<String, Object>> list = handler.importExcel(is, new ImportVo(), "xlsx");
    for (Map<String, Object> map : list) {
      System.out.print(map.get("name"));
    }
  }
}
