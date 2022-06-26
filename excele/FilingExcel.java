package com.spr.hcase.excele;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

/**
 * spr
 *
 * @Description :
 * @Author : JiangQi Luo
 * @Date : 2022/6/21 16:57
 */
@Data
public class FilingExcel {

	@ColumnWidth(20)
	@ExcelProperty("申请人姓名")
	private String applyName;

	@ColumnWidth(20)
	@ExcelProperty("申请人联系方式")
	private String applyPhone;

	@ColumnWidth(20)
	@ExcelProperty("申请人证件类型")
	private String applyCardType;

	@ColumnWidth(20)
	@ExcelProperty("申请人证件号码")
	private String applyCardNo;

	@ColumnWidth(20)
	@ExcelProperty("申请人详细地址")
	private String applyAddress;


	@ColumnWidth(20)
	@ExcelProperty("代理人姓名")
	private String agentName;

	@ColumnWidth(20)
	@ExcelProperty("代理人联系方式")
	private String agentPhone;

	@ColumnWidth(20)
	@ExcelProperty("代理人证件类型")
	private String agentCardType;

	@ColumnWidth(20)
	@ExcelProperty("代理人证件号码")
	private String agentCardNo;

	@ColumnWidth(20)
	@ExcelProperty("代理人详细地址")
	private String agentAddress;

	@ColumnWidth(20)
	@ExcelProperty("ems送达收件人姓名")
	private String emsSendName;

	@ColumnWidth(20)
	@ExcelProperty("ems送达收件人联系方式")
	private String emsSendPhone;

	@ColumnWidth(20)
	@ExcelProperty("ems送达收件人邮箱")
	private String emsSendMail;

	@ColumnWidth(20)
	@ExcelProperty("送达详细地址")
	private String emsSendAddress;

	@ColumnWidth(20)
	@ExcelProperty("被申请人姓名")
	private String respondentName;

	@ColumnWidth(20)
	@ExcelProperty("被申请人联系方式")
	private String respondentPhone;


	@ColumnWidth(20)
	@ExcelProperty("被申请人证件类型")
	private String respondentCardType;

	@ColumnWidth(20)
	@ExcelProperty("被申请人证件号码")
	private String respondentCardNo;

	@ColumnWidth(20)
	@ExcelProperty("被申请人详细地址")
	private String respondentAddress;

	@ColumnWidth(20)
	@ExcelProperty("案件信息")
	private String factReason;

	@ColumnWidth(20)
	@ExcelProperty("案由")
	private String caseReason;

	@ColumnWidth(20)
	@ExcelProperty("仲裁项")
	private String applyItem;

	@ColumnWidth(20)
	@ExcelProperty("详细描述")
	private String applyItemDetail;


	@ColumnWidth(20)
	@ExcelProperty("证据材料名称")
	private String evidenceName;

	@ColumnWidth(20)
	@ExcelProperty("是否为原件")
	private String IsOriginal;

	@ColumnWidth(20)
	@ExcelProperty("证据材料说明")
	private String proveContent;

	@ColumnWidth(20)
	@ExcelProperty("是否接受调节")
	private String isAdjust;

	@ColumnWidth(20)
	@ExcelProperty("备注")
	private String remark;



}
