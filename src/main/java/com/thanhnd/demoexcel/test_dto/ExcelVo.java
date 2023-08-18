package com.thanhnd.demoexcel.test_dto;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.enums.poi.FillPatternTypeEnum;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 10)
// 头字体设置成20
@HeadFontStyle(fontHeightInPoints = 20)
// 内容的背景设置成绿色 IndexedColors.GREEN.getIndex()
@ContentStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 17)
// 内容字体设置成20
@ContentFontStyle(fontHeightInPoints = 13,fontName = "Cantarell")
public class ExcelVo {

	// 字符串的头背景设置成粉红 IndexedColors.PINK.getIndex()
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 14)
	// 字符串的头字体设置成20
	@HeadFontStyle(fontHeightInPoints = 30)
	// 字符串的内容的背景设置成天蓝 IndexedColors.SKY_BLUE.getIndex()
	@ContentStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 40)
	// 字符串的内容字体设置成20
	@ContentFontStyle(fontHeightInPoints = 30)
	@ExcelProperty("Cột số 1")
	private String column1;
	@ExcelProperty("Cột số 2")
	private String column2;
	@ExcelProperty("Cột số 3")
	private String column3;
	private String column4;
	private String column5;
	private String column6;
	private String column7;
	private String column8;
	private String column9;
	private String column10;
	private String column11;
	private String column12;
	private String column13;
	private String column14;
	private String column15;
	private String column16;
	private String column17;
	private String column18;
	private String column19;
	private String column20;
	private String column21;
	private String column22;
	private String column23;
	private String column24;
	private String column25;
	private String column26;
	private String column27;
	private String column28;
	private String column29;
	private String column30;
	private String column31;
	private String column32;
	private String column33;
	private String column34;
	private String column35;
	private String column36;
	private String column37;
	private String column38;
	private String column39;
	private String column40;
	private String column41;
	private String column42;
	private String column43;
	private String column44;
	private String column45;
	private String column46;
	private String column47;
	private String column48;
	private String column49;
	private String column50;
	private String column51;
	private String column52;
	private String column53;
	private String column54;
	private String column55;
	private String column56;
	private String column57;
	private String column58;
	private String column59;
	private String column60;
	private String column61;
	private String column62;
	private String column63;
	private String column64;
	private String column65;
	private String column66;
	private String column67;
	private String column68;
	private String column69;
	private String column70;
	private String column71;
	private String column72;
	private String column73;
	private String column74;
	private String column75;
	private String column76;
	private String column77;
	private String column78;
	private String column79;
	private String column80;
	private String column81;
	private String column82;
	private String column83;
	private String column84;
	private String column85;
	private String column86;
	private String column87;
	private String column88;
	private String column89;
	private String column90;
	private String column91;
	private String column92;
	private String column93;
	private String column94;
	private String column95;
	private String column96;
	private String column97;
	private String column98;
	private String column99;
	private String column100;

}
