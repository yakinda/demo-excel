package com.thanhnd.demoexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.thanhnd.demoexcel.test_dto.ExcelMockData;
import com.thanhnd.demoexcel.test_dto.ExcelUtil;
import com.thanhnd.demoexcel.test_dto.ExcelVo;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.*;
import org.springframework.util.StopWatch;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.zip.GZIPOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@Slf4j
public class TestController {
    @GetMapping("/test")
    public void downLoad(HttpServletResponse response) throws IOException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelMockData mockData = new ExcelMockData();
        List<ExcelVo> dataMoxk = mockData.getExcelData();
        int chunk = 65000;
        List<List<ExcelVo>> listChunk = new ArrayList<>();
        int i = 0;
        while (i < dataMoxk.size()) {
            listChunk.add(dataMoxk.subList(i, Math.min(i + chunk, dataMoxk.size())));
            i += chunk;
        }
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("Kỳ lương tháng 9 năm 2023", StandardCharsets.UTF_8).replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), ExcelVo.class).sheet("sheet1").doWrite(listChunk.get(0));
        stopWatch.stop();
        log.info("Time to process in {} ms", stopWatch.getTotalTimeMillis());
    }

    @GetMapping("/test2")
    public void test2() {
        ExcelMockData mockData = new ExcelMockData();
        List<ExcelVo> dataMoxk = mockData.getExcelData();
        int chunk = 20000;
        List<List<ExcelVo>> listChunk = new ArrayList<>();
        int i = 0;
        while (i < dataMoxk.size()) {
            listChunk.add(dataMoxk.subList(i, Math.min(i + chunk, dataMoxk.size())));
            i += chunk;
        }
        File file = new File("/home/gin/workspace/me/test/demo-excel/src/main/resources/excel/ezExcel.xlsx");
        File filePoi = new File("/home/gin/workspace/me/test/demo-excel/src/main/resources/excel/filePoi.xlsx");
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelWriter excelWriter = EasyExcel.write(file, ExcelVo.class).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
//        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        response.setCharacterEncoding("utf-8");
//        String fileName = URLEncoder.encode("Kỳ lương tháng 9 năm 2023", "UTF-8").replaceAll("\\+", "%20");
//        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        for (List<ExcelVo> data : listChunk) {
            excelWriter.write(data, writeSheet);
        }
        excelWriter.finish();

        stopWatch.stop();
        log.info("Ez excel in {} ms", stopWatch.getTotalTimeMillis());

        try (FileOutputStream fileOutputStream = new FileOutputStream(filePoi)) {
            StopWatch stopWatch1 = new StopWatch();
            stopWatch1.start();
            SXSSFWorkbook workbook = new SXSSFWorkbook(100);
            //workbook.setCompressTempFiles(true);
            SXSSFSheet sheet = workbook.createSheet("sheet1");
            @SuppressWarnings("unchecked")
            Class<ExcelVo> classz = (Class<ExcelVo>) dataMoxk.get(0).getClass();
            Field[] fields = classz.getDeclaredFields();
            int noOfFields = fields.length;
            int rownum = 0;
            Row rowHead = sheet.createRow(rownum);
            for (int o = 0; o < noOfFields; o++) {
                Cell cell = rowHead.createCell(o);
                cell.setCellValue(fields[o].getName());
            }
            for (ExcelVo excelModel : dataMoxk) {
                SXSSFRow row = sheet.createRow(rownum);
                int colnum = 0;
                for (Field field : fields) {
                    String fieldName = field.getName();
                    SXSSFCell cell = row.createCell(colnum);
                    Method method;
                    try {
                        method = classz.getMethod("get" + ExcelUtil.capitalizeInitialLetter(fieldName));
                    } catch (NoSuchMethodException nme) {
                        method = classz.getMethod("get" + fieldName);
                    }
                    Object value = method.invoke(excelModel, (Object[]) null);
                    cell.setCellValue((String) value);
                    colnum++;
                }
                rownum++;
            }
            workbook.write(fileOutputStream);
            workbook.dispose();
            workbook.close();
            stopWatch1.stop();
            log.info("Poi process in {} ms", stopWatch1.getTotalTimeMillis());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @GetMapping("/test-4")
    public ResponseEntity<ByteArrayResource> downLoad2() throws IOException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelMockData mockData = new ExcelMockData();
        List<ExcelVo> dataMoxk = mockData.getExcelData();
        int chunk = 65000;
        List<List<ExcelVo>> listChunk = new ArrayList<>();
        int i = 0;
        while (i < dataMoxk.size()) {
            listChunk.add(dataMoxk.subList(i, Math.min(i + chunk, dataMoxk.size())));
            i += chunk;
        }

        ByteArrayOutputStream excelStream = new ByteArrayOutputStream();
        EasyExcel.write(excelStream, ExcelVo.class).sheet("sheet1").doWrite(listChunk.get(0));

        ByteArrayOutputStream gzippedExcelStream = new ByteArrayOutputStream();
        try (GZIPOutputStream gzipOutputStream = new GZIPOutputStream(gzippedExcelStream)) {
            gzipOutputStream.write(excelStream.toByteArray());
        }

        // Đổi dữ liệu nén thành mảng byte
        byte[] gzippedExcelData = gzippedExcelStream.toByteArray();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentDisposition(ContentDisposition.builder("attachment")
                .filename("excel_file.gz")
                .build());
        headers.set("Content-Encoding", "gzip");
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

//            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        response.setCharacterEncoding("utf-8");
//        String fileName = URLEncoder.encode("Kỳ lương tháng 9 năm 2023", StandardCharsets.UTF_8).replaceAll("\\+", "%20");
//        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
//        EasyExcel.write(response.getOutputStream(), ExcelVo.class).sheet("sheet1").doWrite(listChunk.get(0));
        stopWatch.stop();
        log.info("Time to process in {} ms", stopWatch.getTotalTimeMillis());
        return new ResponseEntity<>(new ByteArrayResource(gzippedExcelData), headers, HttpStatus.OK);

    }


    @GetMapping("/test-3")
    public ResponseEntity<byte[]> downLoad3() throws IOException {

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

        CompletableFuture.allOf(listChunk.stream()
                .map(item -> CompletableFuture.runAsync(() -> {
                            String fiaa = String.format(fileName, atomicInteger.getAndIncrement());
                            listFileName.add(fiaa);
                            EasyExcel
                                    .write(fiaa, ExcelVo.class)
                                    .sheet("Sheet1")
                                    .doWrite(item);
                        }, executor)
                        .thenRun(() -> log.info("{} fill success.", System.currentTimeMillis()))).toArray(CompletableFuture[]::new)).join();

        ByteArrayOutputStream zipByteArrayStream = new ByteArrayOutputStream();
        ZipOutputStream zipOutputStream = new ZipOutputStream(zipByteArrayStream);

        for (int j = 0; j < listFileName.size(); j++) {
            ZipEntry zipEntry = new ZipEntry(String.format("exelfile(%s).xlsx", j));
            zipOutputStream.putNextEntry(zipEntry);
            zipOutputStream.write(Files.readAllBytes(Paths.get(listFileName.get(j))));
            zipOutputStream.closeEntry();
        }

        // Close the ZipOutputStream.
        zipOutputStream.close();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentDisposition(ContentDisposition.builder("attachment")
                .filename("excel_files.zip")
                .build());
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

        // Return the compressed file as a response.
        return new ResponseEntity<>(zipByteArrayStream.toByteArray(), headers, HttpStatus.OK);
    }

}
