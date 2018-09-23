package xin.zhuyao.easyexcelutil.excel;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

/**
 * @Author zy
 * @Date 2018-09-24
 */
public class EasyExcelUtil {
    /**
     * 读取 Excel(多个 sheet)
     *
     * @param excel  文件
     * @param object 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel object) {
        return getReader(excel, object, 0);
    }

    /**
     * 读取某个 sheet 的 Excel
     *
     * @param excel   文件
     * @param object  实体类映射，继承 BaseRowModel 类
     * @param sheetNo sheet 的序号
     *                当前版本中：
     *                XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
     *                XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
     * @return Excel 数据 list
     */

    public static List<Object> readExcel(MultipartFile excel, BaseRowModel object, int sheetNo) {
        return getReader(excel, object, sheetNo);
    }


    /**
     * 读取 Excel(多个 sheet)
     *
     * @param excel  文件
     * @param object 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(File excel, BaseRowModel object) {
        return getReader(excel, object, 0);
    }

    /**
     * 读取某个 sheet 的 Excel
     *
     * @param excel   文件
     * @param object  实体类映射，继承 BaseRowModel 类
     * @param sheetNo sheet 的序号
     *                当前版本中：
     *                XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
     *                XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
     * @return Excel 数据 list
     */

    public static List<Object> readExcel(File excel, BaseRowModel object, int sheetNo) {
        return getReader(excel, object, sheetNo);
    }

    /**
     * 读取 Excel(多个 sheet)
     *
     * @param fileName  文件名
     * @param inputStream 输入流
     * @param object 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(String fileName, InputStream inputStream, BaseRowModel object) {
        return getReader(fileName, inputStream, object, 0);
    }

    /**
     * 读取某个 sheet 的 Excel
     *
     * @param fileName   文件名
     * @param object  实体类映射，继承 BaseRowModel 类
     * @param inputStream  输入流
     * @param sheetNo sheet 的序号
     *                当前版本中：
     *                XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
     *                XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(String fileName, InputStream inputStream, BaseRowModel object, int sheetNo) {
        return getReader(fileName, inputStream, object, sheetNo);
    }




    /**
     * 导出 Excel ：一个或多个 sheet，带表头
     *
     * @param response  HttpServletResponse
     * @param fileName  导出的文件名
     */
    public static EasyExcelWriterFactroy writeExcelWithSheets(HttpServletResponse response, String fileName) throws IOException {
        //创建本地文件
        String filePath = fileName + ".xlsx";
        File dbfFile = new File(filePath);
        if (!dbfFile.exists() || dbfFile.isDirectory()) {
            dbfFile.createNewFile();
        }
        fileName = new String(filePath.getBytes(), "ISO-8859-1");
        response.addHeader("Content-Disposition", "filename=" + fileName);
        OutputStream out = response.getOutputStream();
        EasyExcelWriterFactroy writer = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
        return writer;
    }

    /**
     * 导出 Excel ：一个或多个 sheet，带表头
     *
     * @param fileName  导出的文件名
     */
    public static EasyExcelWriterFactroy writeExcelWithSheets(String fileName) throws IOException {
        OutputStream out = new FileOutputStream(fileName);
        EasyExcelWriterFactroy writer = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
        return writer;
    }


    /**
     * 返回 Excel 数据 list
     *
     * @param excel  需要解析的 Excel 文件
     * @param object 实体类映射，继承 BaseRowModel 类
     */
    private static List<Object> getReader(MultipartFile excel, BaseRowModel object, int sheetNo) {
        String fileName = excel.getOriginalFilename();
        InputStream inputStream;
        try {
            inputStream = excel.getInputStream();
            return getReader(fileName, inputStream, object, sheetNo);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }


    /**
     * 返回 Excel 数据 list
     *
     * @param excel   需要解析的 Excel 文件
     * @param object 实体类映射，继承 BaseRowModel 类
     */
    private static  List<Object> getReader(File excel, BaseRowModel object, int sheetNo) {
        String fileName = excel.getName();
        InputStream inputStream;
        try {
            inputStream = new FileInputStream(excel);
            return getReader(fileName, inputStream, object, sheetNo);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     *
     */
    private static List<Object> getReader(String fileName, InputStream inputStream, BaseRowModel object, int sheetNo) {
        if (fileName == null || (!fileName.toLowerCase().endsWith(".xls") && !fileName.toLowerCase().endsWith(".xlsx"))) {
            throw new EasyExcelException("文件格式错误！");
        }
        ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.XLSX;
        if (fileName.toLowerCase().endsWith(".xls")) {
            excelTypeEnum = ExcelTypeEnum.XLS;
        }
        EasyExcelListener easyExcelListener = new EasyExcelListener();
        ExcelReader reader = new ExcelReader(inputStream, excelTypeEnum, null, easyExcelListener);
        if (reader == null) {
            return null;
        }
        if (sheetNo == 0) {
            for (Sheet sheet : reader.getSheets()) {
                if (object != null) {
                    sheet.setClazz(object.getClass());
                }
                reader.read(sheet);
            }
        }else {
            Sheet sheet = new Sheet(sheetNo);
            sheet.setClazz(object.getClass());
            reader.read(sheet);
        }
        return easyExcelListener.getDatas();
    }
}
