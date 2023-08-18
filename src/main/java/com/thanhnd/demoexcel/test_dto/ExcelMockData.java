package com.thanhnd.demoexcel.test_dto;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.List;

public class ExcelMockData {

	private List<ExcelVo> excelData;

	public ExcelMockData() {
	}

	public List<ExcelVo> getExcelData() {
		if (excelData == null) {
			populateExcelData();
		}
		return excelData;
	}

	public void setExcelData(List<ExcelVo> excelData) {
		this.excelData = excelData;
	}

	private void populateExcelData() {
		excelData = new ArrayList<>();
		Class<ExcelVo> classz = (Class<ExcelVo>) ExcelVo.class;
		Method[] methods = classz.getMethods();
		ExcelVo model = new ExcelVo();
		for (Method method : methods) {
			String methodName = method.getName();
			if (methodName.startsWith("set")) {
				try {
					method.invoke(model, getRandomString());
				} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
					e.printStackTrace();
				}
			}
		}
		for (int i = 0; i < 800000; i++) {
			excelData.add(model);
		}
	}

	private String getRandomString() {
		SecureRandom random = new SecureRandom();
		return new BigInteger(130, random).toString(32);
	}

}
