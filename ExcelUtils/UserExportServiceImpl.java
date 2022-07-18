package com.mukun.user.impl.self;

import com.datedu.base.model.vo.UserInfoVO;
import com.datedu.base.service.BaseLoginService;
import com.mukun.base.dto.SortDto;
import com.mukun.base.service.SchoolService;
import com.mukun.base.service.SortService;
import com.mukun.component.utils.ApiResult;
import com.mukun.interceptor.interceptor.UserThreadLocal;
import com.mukun.interceptor.provider.UserInfo;
import com.mukun.user.config.insideService.UserExportService;
import com.mukun.user.config.utils.CommonUtil;
import com.mukun.user.config.utils.ExcelUtils;
import com.mukun.user.dao.mysql.TpUserInfoMapper;
import com.mukun.user.dto.StudentDto;
import com.mukun.user.dto.TeacherDto;
import com.mukun.user.service.StudentSelfService;
import mukun.service.tokenstore.provider.TokenStore;
import org.apache.commons.lang3.StringUtils;
import org.apache.dubbo.config.annotation.DubboReference;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class UserExportServiceImpl implements UserExportService {

    @DubboReference(check = false)
    SortService sortService;

    @DubboReference(check = false)
    SchoolService schoolService;

    @Autowired
    TokenStore tokenStore;

    @Autowired
    TpUserInfoMapper tpUserInfoMapper;

    @Autowired
    StudentSelfService studentSelfService;
    @DubboReference(check = false, registry = "pj")
    BaseLoginService baseLoginService;

    @Override
    public ApiResult exportTea(HttpServletResponse response, HttpServletRequest request, Integer schoolType, String access_token) {
//        String schoolId = getUserInfoByToken(tokenStore, access_token).getSchoolId();//getSchoolId(access_token);
//        if (StringUtils.isEmpty(schoolId)) {
//            return ApiResult.fail(ErrorConstant.NO_GET_LOGIN);
//        }
        UserInfo userInfo = UserThreadLocal.get();
        if (userInfo == null) {
            return ApiResult.error("用户信息不存在");
        }
        String subId = null;
        if (userInfo.getRoleTypes().contains("5")) {
            subId = userInfo.getUserSubId();
        }
        //String schoolId = UserThreadLocal.get().getSchoolId();
        schoolType = schoolType == null ? schoolService.getSchoolType(userInfo.getSchoolId()) : schoolType;
        List<TeacherDto> teaList = tpUserInfoMapper.getListBySchoolId(null, userInfo.getSchoolId(), null, new ArrayList<>(), null, null, null, 0, null);
        if (teaList.size() == 0) {
            return ApiResult.error("没有用户需要导出");
        }
        // List<String> baseUserIds = teaList.stream().map(TeacherDto::getBaseUserId).collect(Collectors.toList());
        //List<PjMapUserDto> pjMapUserDtos = studentSelfService.getPiUserList(schoolId, null, baseUserIds, 1, baseUserIds.size(), 2,1);
        String[] title = {"用户名", "工号", "姓名", "创建时间", "院系", "手机号", "是否建课", "是否禁用"};
        String[][] content = new String[teaList.size()][7];
        List<SortDto> sortList = new ArrayList<>();
        if (schoolType == 4) {  //企业
            title = new String[]{"用户名", "工号", "姓名", "创建时间", "组织架构", "手机号", "是否建课", "是否禁用"};
            List<String> sortIdList = CommonUtil.getIdListByNodePath(teaList.stream().filter(q -> StringUtils.isNotEmpty(q.getDepartmentId())).map(TeacherDto::getDepartmentId).collect(Collectors.toList()));
            sortList = sortService.getSortList(userInfo.getSchoolId(), sortIdList);
        } else if (schoolType == 7 || schoolType == 8 || schoolType == 9) {
            title = new String[]{"用户名", "工号", "姓名", "创建时间", "手机号", "是否建课", "是否禁用"};
        }
        for (int i = 0; i < teaList.size(); i++) {
            content[i] = new String[title.length];
            TeacherDto obj = teaList.get(i);
//            PjMapUserDto pjMapUserDto = pjMapUserDtos.stream().filter(q -> q.getId().equals(obj.getBaseUserId())).findFirst().orElse(null);
//            if (pjMapUserDto == null) {
//                continue;
//            }
            content[i][0] = obj.getUserName();
            content[i][1] = obj.getEmployeeNumber();
            content[i][2] = obj.getDisplayName();
            content[i][3] = obj.getCreateTime();
            if (schoolType == 7 || schoolType == 8 || schoolType == 9) {
                content[i][4] = obj.getMobile();
                content[i][5] = obj.isTeacherSpace() ? "是" : "否";
                content[i][6] = !obj.isIsValid() ? "是" : "否";
            } else {
                String sortName = "";
                if (schoolType == 4) {
                    if (StringUtils.isNotEmpty(obj.getDepartmentId()) && sortList != null && sortList.size() > 0) {
                        sortName = CommonUtil.getSortPathName(obj.getDepartmentId(), sortList);
                    }
                } else {
                    sortName = obj.getDepartmentName();
                }
                content[i][4] = sortName;
                content[i][5] = obj.getMobile();
                content[i][6] = obj.isTeacherSpace() ? "是" : "否";
                content[i][7] = !obj.isIsValid() ? "是" : "否";
            }
        }
        String workTitle = "教师";
        try {
            // 解决文件乱码
            final String userAgent = request.getHeader("user-agent");
            if (userAgent != null && userAgent.contains("Firefox")) {
                workTitle = new String(workTitle.getBytes(), StandardCharsets.ISO_8859_1);
            } else {
                workTitle = URLEncoder.encode(workTitle, "UTF8");
            }
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        if (StringUtils.isEmpty(workTitle)) {
            workTitle = "teacher";
        }
        //创建HSSFWorkbook
        ExcelUtils.getHSSFWorkbook(response, workTitle, title, content);
        return null;
    }

    @Override
    public ApiResult exportStu(HttpServletResponse response, HttpServletRequest request, Integer schoolType, String access_token) {
//        String schoolId = getUserInfoByToken(tokenStore, access_token).getSchoolId();// getSchoolId(access_token);
//        if (StringUtils.isEmpty(schoolId)) {
//            return ApiResult.fail(ErrorConstant.NO_GET_LOGIN);
//        }
        UserInfo userInfo = UserThreadLocal.get();
        if (userInfo == null) {
            return ApiResult.error("用户信息不存在");
        }
        String subId = null;
        if (userInfo.getRoleTypes().contains("5")) {
            subId = userInfo.getUserSubId();
        }
        schoolType = schoolType == null ? schoolService.getSchoolType(userInfo.getSchoolId()) : schoolType;
        String[] title = {"用户名", "学号", "姓名", "创建时间", "院系", "专业", "年级", "行政班", "手机号", "是否禁用"};
        List<StudentDto> stuList = tpUserInfoMapper.getListBySchoolIdForSchool(subId, userInfo.getSchoolId(), null,
                null, null, null, null, new ArrayList<>(), null);
        if (stuList.size() == 0) {
            return ApiResult.error("没有用户需要导出");
        }
        List<UserInfoVO> userInfoDTOS = new ArrayList<>();
        if (schoolType == 7 || schoolType == 8 || schoolType == 9) {
            title = new String[]{"用户名", "学号", "姓名", "创建时间", "行政班", "手机号", "是否禁用"};
            List<String> ids = stuList.stream().map(StudentDto::getBaseUserId).collect(Collectors.toList());
            userInfoDTOS = (List<UserInfoVO>) baseLoginService.getUserInfoListSupportMultiSchool(ids, null, true).getData();
        }
        List<String> baseUserIds = stuList.stream().map(StudentDto::getBaseUserId).collect(Collectors.toList());
        //List<PjMapUserDto> pjMapUserDtos = studentSelfService.getPiUserList(userInfo.getSchoolId(), null, baseUserIds, 1, baseUserIds.size(), 1,1);
        String[][] content = new String[stuList.size()][9];
        for (int i = 0; i < stuList.size(); i++) {
            content[i] = new String[title.length];
            StudentDto obj = stuList.get(i);
//            PjMapUserDto pjMapUserDto = pjMapUserDtos.stream().filter(q -> q.getId().equals(obj.getBaseUserId())).findFirst().orElse(null);
//            if (pjMapUserDto == null) {
//                continue;
//            }
            content[i][0] = obj.getUserName();
            content[i][1] = obj.getEmployeeNumber();
            content[i][2] = obj.getDisplayName();
            content[i][3] = obj.getCreateTime();
            if (schoolType == 7 || schoolType == 8 || schoolType == 9) {
                if (userInfoDTOS.size() > 0) {
                    UserInfoVO info = userInfoDTOS.stream().filter(q -> q.getId().equals(obj.getBaseUserId())).findFirst().orElse(null);
                    if (info != null) {
                        content[i][4] = info.getStudentAdminClassName();
                    }
                }
                content[i][5] = obj.getMobile();
                content[i][6] = obj.isIsValid() ? "否" : "是";
            } else {
                content[i][4] = obj.getDepartmentName();
                content[i][5] = obj.getMajorName();
                content[i][6] = obj.getGradeName();
                content[i][7] = obj.getClassName();
                content[i][8] = obj.getMobile();
                content[i][9] = obj.isIsValid() ? "否" : "是";
            }
        }
        String workTitle = "学生";
        try {
            // 解决文件乱码
            final String userAgent = request.getHeader("user-agent");
            if (userAgent != null && userAgent.contains("Firefox")) {
                workTitle = new String(workTitle.getBytes(), StandardCharsets.ISO_8859_1);
            } else {
                workTitle = URLEncoder.encode(workTitle, "UTF8");
            }
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        if (StringUtils.isEmpty(workTitle)) {
            workTitle = "student";
        }
        //创建HSSFWorkbook
        ExcelUtils.getHSSFWorkbook(response, workTitle, title, content);
        return null;
    }

}
