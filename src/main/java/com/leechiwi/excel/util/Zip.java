package com.leechiwi.excel.util;


import org.apache.commons.collections.CollectionUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Zip {
    public static void zip(OutputStream out, List<byte[]> streamList, List<String> fileNames,String ext) throws RuntimeException {
        if(CollectionUtils.isEmpty(streamList)){
            throw new RuntimeException("压缩错误,没有被压缩的文件");
        }
        if(CollectionUtils.isEmpty(fileNames)){
            fileNames = new ArrayList<>();
            for(int i=0,len=streamList.size();i<len;i++){
                StringBuilder sb=new StringBuilder();
                fileNames.add(sb.append(i).append(ext).toString());
            }
        }
        if(streamList.size()!=fileNames.size()){
            throw new RuntimeException("压缩错误,被压缩的文件数量和压缩文件的名称数量不相等");
        }
        ZipOutputStream zipOut = new ZipOutputStream(out);
        int index=0;
        try {
            for (byte[] bytes : streamList) {
                // 创建一个ZipEntry
                zipOut.putNextEntry(new ZipEntry(fileNames.get(index)+".xlsx"));
                // 将源文件的字节内容，写入zip压缩包
                zipOut.write(bytes);
                // 结束当前zipEntry
                zipOut.closeEntry();
            }
            zipOut.finish();
            zipOut.close();
        } catch (IOException e) {
            throw new RuntimeException("zip error");
        }
    }
}
