package com.example.excel.pojo;

/**
 * EXCEL表头以及HBase列限定名的对照字典表.
 */
public enum TableHeadEnum {
	/** 文件表头限定对照 */
	HEAD_UNKNOWN("未知类型", "UNKNOWN"),

	SCORE_HEAD_1("考号", "ENUM"),
	SCORE_HEAD_2("学籍号", "UNKNOWN"),
	SCORE_HEAD_3("班级", "UNAME"),
	SCORE_HEAD_4("姓名", "STU_NAME"),
	SCORE_HEAD_5("学生属性", "UNKNOWN"),
	SCORE_HEAD_6("语文得分", "CN_SCORE"),
	SCORE_HEAD_7("语文校次", "CN_ORANK"),
	SCORE_HEAD_9("语文班次", "CN_URANK"),
	SCORE_HEAD_11("数学得分", "MATH_SCORE"),
	SCORE_HEAD_12("数学校次", "MATH_ORANK"),
	SCORE_HEAD_14("数学班次", "MATH_URANK"),
	SCORE_HEAD_16("英语得分", "EN_SCORE"),
	SCORE_HEAD_17("英语校次", "EN_ORANK"),
	SCORE_HEAD_19("英语班次", "EN_URANK"),
	SCORE_HEAD_21("物理得分", "Q_SCORE"),
	SCORE_HEAD_22("物理校次", "Q_ORANK"),
	SCORE_HEAD_24("物理班次", "Q_URANK"),
	SCORE_HEAD_26("化学得分", "CHEM_SCORE"),
	SCORE_HEAD_27("化学校次", "CHEM_ORANK"),
	SCORE_HEAD_29("化学班次", "CHEM_URANK"),
	SCORE_HEAD_30("生物得分", "BOIL_SCORE"),
	SCORE_HEAD_31("生物校次", "BOIL_ORANK"),
	SCORE_HEAD_32("生物班次", "BOIL_URANK"),
	SCORE_HEAD_33("政治得分", "POL_SCORE"),
	SCORE_HEAD_34("政治校次", "POL_ORANK"),
	SCORE_HEAD_35("政治班次", "POL_URANK"),
	SCORE_HEAD_36("历史得分", "HIST_SCORE"),
	SCORE_HEAD_37("历史校次", "HIST_ORANK"),
	SCORE_HEAD_40("历史班次", "HIST_URANK"),
	SCORE_HEAD_41("地理得分", "GEO_SCORE"),
	SCORE_HEAD_42("地理校次", "GEO_ORANK"),
	SCORE_HEAD_43("地理班次", "GEO_URANK"),
	SCORE_HEAD_44("总分得分", "T_SCORE"),
	SCORE_HEAD_45("总分校次", "T_ORANK"),
	SCORE_HEAD_46("总分班次", "T_URANK");


	/** EXCEL表中的表头中文名称 */
	private String headName;

	/** HBase表列限定名 */
	private String columnQualifier;

	TableHeadEnum(String headName, String columnQualifier) {
		this.headName = headName;
		this.columnQualifier = columnQualifier;
	}

	/**
	 * 根据表头的中文名称获取HBase表中的列限定名
	 * @param headName
	 * @return
	 */
	public static TableHeadEnum getByHeadName(String headName) {
		for (TableHeadEnum tableHeadEnum : TableHeadEnum.values()) {
			if (tableHeadEnum.getHeadName().equals(headName)) {
				return tableHeadEnum;
			}
		}
		return HEAD_UNKNOWN;
	}

	public String getHeadName() {
		return headName;
	}

	public void setHeadName(String headName) {
		this.headName = headName;
	}

	public String getColumnQualifier() {
		return columnQualifier;
	}

	public void setColumnQualifier(String columnQualifier) {
		this.columnQualifier = columnQualifier;
	}
}
