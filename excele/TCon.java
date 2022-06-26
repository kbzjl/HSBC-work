package com.spr.hcase.excele;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.ObjectUtil;
import com.alibaba.excel.EasyExcel;
import com.spr.common.utils.R;
import com.spr.common.utils.Rcode;
import com.spr.common.utils.RedisJwt;
import com.spr.hcase.constant.CasesConstant;
import com.spr.hcase.entity.*;
import com.spr.hcase.service.*;
import com.spr.hcase.utils.ExportExcelUtil;
import com.spr.hcase.utils.SM4Utils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/servicehcase/et")
public class TCon {

    @Autowired
    private CasesService casesService;

    @Autowired
    private PersonService personService;

    @Autowired
    private PersonMailService personMailService;

    @Autowired
    private CaseApplyItemService caseApplyItemService;

    @Autowired
    private EvidenceService evidenceService;

    @Autowired
    private RedisJwt redisJwt;


    @Value("${init.key}")
    private String key;

    @PostMapping("/im")
    public R importExcel(HttpServletRequest request, @RequestParam("file") MultipartFile file) throws IOException {
        Map<String, String> userMap = redisJwt.getUserByJwtToken(request);
        if (CollectionUtil.isEmpty(userMap)) {
            return new R(false, Rcode.TOKEN_EFFICACY, "登录信息已失效，请重新登录！", null);
        }
        String orgId = userMap.get("orgId");
        InputStream inputStream = file.getInputStream();
        // excel 读取的集合
        List<FilingExcel> allList = EasyExcel.read(inputStream)
                .head(FilingExcel.class)
                // 设置sheet,默认读取第一个
                .sheet()
                // 设置标题所在行数
                .headRowNumber(1)
                .doReadSync();
        for (FilingExcel filingExcel : allList) {
            if (ObjectUtil.isNotNull(filingExcel)){
                if (StringUtils.isEmpty(filingExcel.getApplyName())) {
                    return R.error().message("缺少申请人姓名");
                }
                if (StringUtils.isEmpty(filingExcel.getApplyPhone())) {
                    return R.error().message("缺少申请人联系方式");
                }
                if (StringUtils.isEmpty(filingExcel.getApplyCardType())) {
                    return R.error().message("缺少申请人证件类型");
                }
                if (StringUtils.isEmpty(filingExcel.getApplyCardNo())) {
                    return R.error().message("缺少申请人证件号码");
                }
                if (StringUtils.isEmpty(filingExcel.getEmsSendName())) {
                    return R.error().message("缺少收件人姓名");
                }
                if (StringUtils.isEmpty(filingExcel.getEmsSendPhone())) {
                    return R.error().message("缺少收件人联系方式");
                }
                if (StringUtils.isEmpty(filingExcel.getEmsSendMail())) {
                    return R.error().message("缺少收件人邮箱");
                }
                if (StringUtils.isEmpty(filingExcel.getRespondentName())) {
                    return R.error().message("缺少被申请人姓名");
                }
                if (StringUtils.isEmpty(filingExcel.getRespondentPhone())) {
                    return R.error().message("缺少被申请人联系方式");
                }
                if (StringUtils.isEmpty(filingExcel.getRespondentCardType())) {
                    return R.error().message("缺少被申请人证件类型");
                }
                if (StringUtils.isEmpty(filingExcel.getRespondentCardNo())) {
                    return R.error().message("缺少被申请人证件号码");
                }
            }
        }

        for (FilingExcel filingExcel : allList) {
            Cases codeAndId = casesService.getCodeAndId(orgId, null);
            // 添加申请人
            Person applyPerson = new Person().setCaseId(codeAndId.getId())
                    .setPersonNature(CasesConstant.APPLY)
                    .setName(filingExcel.getApplyName())
                    .setCardNo(SM4Utils.encrypt(filingExcel.getApplyCardNo(), key))
                    .setPhone(filingExcel.getApplyPhone())
                    .setAddress(filingExcel.getApplyAddress())
                    .setOrgId(orgId);
            personService.save(applyPerson);
            // 添加代理人
            Person agentPerson = new Person().setCaseId(codeAndId.getId())
                    .setPersonNature(CasesConstant.AGENT)
                    .setName(filingExcel.getAgentName())
                    .setCardNo(SM4Utils.encrypt(filingExcel.getAgentCardNo(), key))
                    .setPhone(filingExcel.getAgentPhone())
                    .setAddress(filingExcel.getAgentAddress())
                    .setOrgId(orgId);
            personService.save(agentPerson);
            // 添加被申请人
            Person respondentPerson = new Person().setCaseId(codeAndId.getId())
                    .setPersonNature(CasesConstant.RESPONDENT)
                    .setName(filingExcel.getRespondentName())
                    .setPhone(filingExcel.getRespondentPhone())
                    .setCardNo(SM4Utils.encrypt(filingExcel.getRespondentCardNo(), key))
                    .setAddress(filingExcel.getRespondentAddress())
                    .setOrgId(orgId);
            personService.save(respondentPerson);
            // 添加收件人
            PersonMail personMail = new PersonMail().setCaseId(codeAndId.getId())
                    .setReceiver(filingExcel.getEmsSendName())
                    .setPhone(filingExcel.getEmsSendPhone())
                    .setMailbox(filingExcel.getEmsSendMail())
                    .setMailAddress(filingExcel.getEmsSendAddress());
            personMailService.save(personMail);
            // 添加请求项
            CaseApplyItem caseApplyItem = new CaseApplyItem().setCaseId(codeAndId.getId())
                    .setApplyItem(filingExcel.getApplyItem())
                    .setApplyItemDetail(filingExcel.getApplyItemDetail());
            caseApplyItemService.save(caseApplyItem);
            // 添加证据材料
            Evidence evidence = new Evidence().setCaseId(codeAndId.getId())
                    .setEvidenceName(filingExcel.getEvidenceName())
                    .setProveContent(filingExcel.getProveContent())
                    .setIsOriginal(StringUtils.equals(filingExcel.getIsOriginal(), "是") ? "1":"2");
            evidenceService.save(evidence);
            // 修改案件
            Cases cases = new Cases().setId(codeAndId.getId())
                    .setApplyId(applyPerson.getId())
                    .setApplyName(applyPerson.getName())
                    .setApplyCard(applyPerson.getCardNo())
                    .setAgentId(agentPerson.getId())
                    .setAgentName(agentPerson.getName())
                    .setAgentCard(agentPerson.getCardNo())
                    .setRespondentId(respondentPerson.getId())
                    .setRespondentName(respondentPerson.getName())
                    .setRespondentCard(respondentPerson.getCardNo())
                    .setCaseStatus(CasesConstant.WAIT_ESTABLISH_CASE)
                    .setFactReason(filingExcel.getFactReason())
                    .setCaseReason(filingExcel.getCaseReason())
                    .setMemo1(filingExcel.getRemark());
            casesService.updateById(cases);
        }
        return R.ok();

    }

    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletResponse response) {
        String title = "2021案件统计报表.xlsx";
        String[] headers = {"序号","仲裁委名称","受理案件总数","标的总额","网上仲裁案件数"};
        List<CaseStatistics> cases = Arrays.asList(
                new CaseStatistics("1", "2", "3", "4", "5"),
                new CaseStatistics("2", "2", "3", "4", "5"),
                new CaseStatistics("3", "2", "3", "4", "5")
        );
        ExportExcelUtil exportExcelUtil = new ExportExcelUtil<CaseStatistics>();
        exportExcelUtil.exportExcel(title,headers,cases,response);
    }


}