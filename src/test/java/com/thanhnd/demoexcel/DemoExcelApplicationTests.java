package com.thanhnd.demoexcel;

import com.alibaba.excel.EasyExcel;
import com.thanhnd.demoexcel.test_dto.ExcelMockData;
import com.thanhnd.demoexcel.test_dto.ExcelVo;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@SpringBootTest
@Slf4j
class DemoExcelApplicationTests {

    @Test
    void contextLoads() {
        ExcelMockData mockData = new ExcelMockData();
        List<ExcelVo> data = mockData.getExcelData();
        int chunk = 65000;
        List<List<ExcelVo>> listChunk = new ArrayList<>();
        int i = 0;
        while (i < data.size()) {
            listChunk.add(data.subList(i, Math.min(i + chunk, data.size())));
            i += chunk;
        }
        String fileName = "/home/gin/workspace/me/test/demo-excel/src/main/resources/excel/" + System.currentTimeMillis() + "(%s)" + ".xlsx";
        Executor executor = Executors.newFixedThreadPool(4);
        AtomicInteger atomicInteger = new AtomicInteger(0);

        List<String> listFileName = new CopyOnWriteArrayList<>();
        List<CompletableFuture<Void>> futureList = listChunk.stream()
                .map(item -> CompletableFuture.runAsync(() -> {
                            String fiaa = String.format(fileName, atomicInteger.getAndIncrement());
                            listFileName.add(fiaa);
                            EasyExcel
                                    .write(fiaa, ExcelVo.class)
                                    .sheet("Sheet1")
                                    .doWrite(item);
                        }, executor)
                        .thenRun(() -> log.info("{} fill success.", System.currentTimeMillis())))
                .collect(Collectors.toList());

        CompletableFuture.allOf(futureList.toArray(new CompletableFuture[0])).join();
    }

}
