package xin.zhuyao.easyexcelutil.test;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import xin.zhuyao.easyexcelutil.excel.EasyExcelUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @Author zy
 * @Date 2018-09-24
 */
@RestController
public class ReadController {
    /**
     * 读取 Excel（允许多个 sheet）
     */
    @RequestMapping(value = "readExcel", method = RequestMethod.POST)
    public Object readExcel(MultipartFile excel) {
        return EasyExcelUtil.readExcel(excel, new ImportInfo());
    }


//    public static void main(String[] args) {
//        List<Object> objects = EasyExcelUtil.readExcel(new File("C:\\Users\\zhuyao\\Downloads\\一个 Excel 文件 (1).xlsx"), new ImportInfo());
//        System.out.println(objects.size());
//    }

    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName, new ExportInfo());
    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }

    public static void main(String[] args) throws IOException {
        List<ExportInfo> list = getList();
        String sheetName1 = "第一个 sheet";
        EasyExcelUtil.writeExcelWithSheets("D:\\test.xlsx")
                .write(list,sheetName1,new ExportInfo())
                .finish();
    }


    private static List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("howie");
        model1.setAge("19");
        model1.setAddress("123456789");
        model1.setEmail("123456789@gmail.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("harry");
        model2.setAge("20");
        model2.setAddress("198752233");
        model2.setEmail("198752233@gmail.com");
        list.add(model2);
        return list;
    }

}
