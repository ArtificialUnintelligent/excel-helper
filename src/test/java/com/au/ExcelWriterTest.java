package com.au;

import com.au.model.ExcelModel;
import com.au.resolver.ExcelWriter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author:artificialunintelligent
 * @Date:2019-07-05
 * @Time:17:59
 */
public class ExcelWriterTest {

    @Test
    public void testCreate() throws Exception {
        List<ExcelModel> models = new ArrayList<>();
        final ExcelModel model = new ExcelModel();
        model.setRemark("aaa");
        model.setName("1");
        model.setColor("适合");
        model.setAge("1");
        model.setKind("黄种人");
        model.setSex("男性");
        model.setHeight("1");
        model.setWeight("1");
        model.setJob("工程师");
        final ExcelModel model1 = new ExcelModel();
        model1.setName("2");
        model1.setColor("适合");
        model1.setAge("2");
        model1.setKind("黄种人");
        model1.setSex("男性");
        model1.setHeight("2");
        model1.setWeight("2");
        model1.setJob("工程师");
        final ExcelModel model2 = new ExcelModel();
        model2.setRemark("bbb");
        model2.setName("3");
        model2.setColor("适合");
        model2.setAge("3");
        model2.setKind("种人");
        model2.setSex("男性");
        model2.setHeight("3");
        model2.setWeight("3");
        model2.setJob("工程");
        final ExcelModel model3 = new ExcelModel();
        model3.setName("4");
        model3.setColor("适合");
        model3.setAge("4");
        model3.setKind("黄种人");
        model3.setSex("男性");
        model3.setHeight("4");
        model3.setWeight("4");
        model3.setJob("工程师");
        final ExcelModel model4 = new ExcelModel();
        model4.setName("5");
        model4.setColor("适合");
        model4.setAge("5");
        model4.setKind("黄人");
        model4.setSex("男性");
        model4.setHeight("5");
        model4.setWeight("5");
        model4.setJob("工程师");
        final ExcelModel model5 = new ExcelModel();
        model5.setRemark("aaa1");
        model5.setName("6");
        model5.setColor("适合");
        model5.setAge("6");
        model5.setKind("黄种人");
        model5.setSex("男性");
        model5.setHeight("6");
        model5.setWeight("6");
        model5.setJob("工程师");
        models.add(model);
        models.add(model1);
        models.add(model2);
        models.add(model3);
        models.add(model4);
        models.add(model5);

        ExcelWriter excelWriter = new ExcelWriter();
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(2);
        int i = 0;
        for (int j = 0; j < 3; j++) {
            List<ExcelModel> excelModels = new ArrayList<>();
            excelModels.add(models.get(i++));
            excelModels.add(models.get(i++));
            excelWriter.create(excelModels, sxssfWorkbook, j + 1, 2);
        }

        FileOutputStream out = new FileOutputStream("result.xls");
        sxssfWorkbook.write(out);
        out.close();
    }
}
