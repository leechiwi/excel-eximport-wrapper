package com.leechiwi.excel.mapper;

import com.leechiwi.excel.pojo.Test;
import org.springframework.stereotype.Repository;

@Repository
public class TestMapper {
    public void insert(Test t){
        System.out.println(t.getName());
    }
}
