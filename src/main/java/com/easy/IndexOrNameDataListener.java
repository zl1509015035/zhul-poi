package com.easy;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * @author zhul
 */
public class IndexOrNameDataListener extends AnalysisEventListener<IndexOrNameDate> {
    private static final Logger LOGGER = LoggerFactory.getLogger(IndexOrNameDate.class);

    private static final int BATCH_COUNT = 5;
    List<IndexOrNameDate> list = new ArrayList<>();

    private IndexOrNameDao dao;

    public IndexOrNameDataListener() {
        dao = new IndexOrNameDao();
    }

    public IndexOrNameDataListener(IndexOrNameDao dao) {
        this.dao = dao;
    }

    @Override
    public void invoke(IndexOrNameDate data, AnalysisContext context) {
        LOGGER.info("{}", JSON.toJSONString(data));
        String str = data.getString();
        Double d = data.getDoubleData();
        System.out.println(str+"====="+d);
        list.add(data);
        if (list.size() >= BATCH_COUNT) {
            saveData();
            list.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        saveData();
        LOGGER.info("所有数据解析完成！");
    }

    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        dao.save(list);
        LOGGER.info("存储数据库成功！");
    }
}
