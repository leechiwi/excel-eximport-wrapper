package com.leechiwi.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

public class MyReadListener<T> extends AnalysisEventListener<T> {

    private static final Logger LOGGER = LoggerFactory.getLogger(MyReadListener.class);

    //private final int BATCH_SAVE_NUM = 5;
    private List<T> list = new ArrayList<>();
    private Consumer<T> consumer;
    public MyReadListener(){

    }
    public MyReadListener(Consumer<T> consumer){
        this.consumer=consumer;
    }
    @Override
    public void invoke(T t, AnalysisContext context) {
        list.add(t);
        if(Objects.nonNull(this.consumer)){
           this.consumer.accept(t);
        }
        /*if (students.size() % BATCH_SAVE_NUM == 0) {
            studentService.save(students);
            students.clear();
        }*/
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }

    public List<T> getList() {
        return list;
    }

    public void setList(List<T> list) {
        this.list = list;
    }
}

