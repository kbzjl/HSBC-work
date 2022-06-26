package com.spr.hcase.excele;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.io.Serializable;

/**
 * spr
 *
 * @Description :
 * @Author : JiangQi Luo
 * @Date : 2022/6/24 16:54
 */
@Data
public class CaseStatistics {

	private String num;

	private String orgName;

	private String casesAcceptedSum;

	private String totalAmount;

	private String onlineCasesSum;

	public CaseStatistics(String num, String orgName, String casesAcceptedSum, String totalAmount, String onlineCasesSum) {
		this.num = num;
		this.orgName = orgName;
		this.casesAcceptedSum = casesAcceptedSum;
		this.totalAmount = totalAmount;
		this.onlineCasesSum = onlineCasesSum;
	}
}
