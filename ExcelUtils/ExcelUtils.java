package com.mukun.user.config.utils;

import com.mukun.base.dto.RedisDto;
import com.mukun.component.utils.CheckDataUtils;
import com.mukun.user.dto.ExcelCheckUserDto;
import com.mukun.user.dto.ExcelErrorDto;
import com.mukun.user.dto.ExcelResultDto;
import com.mukun.user.dto.ExcelUserDto;
import com.mukun.user.selfDto.UserInfoDto;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelUtils {
    /**
     * 默认行数（须知、表头）
     */
    public static int DefaultNumberRows = 2;

    /**
     * 导入最大条数
     */
    public static int MaxDefaultNumberRows = 500;

    /**
     * excel转data
     *
     * @param inputStream  文件流
     * @param fileName     文件名称
     * @param userType     用户类型  1，学生  2，教师
     * @param redisDtoList 班级/院系/组织架构集合
     * @param schoolType   学校类型  4，企业  其他学校
     * @param schoolId     学校Id
     * @return com.mukun.user.dto.ExcelResultDto
     * @Description
     * @Author xzli
     * @CreateDate 2021/6/7 15:10
     */
    public static ExcelResultDto getExcelBody(InputStream inputStream, String fileName, int userType, List<RedisDto> redisDtoList,
                                               int schoolType, String schoolId,  List<UserInfoDto> userInfoDtos) {
        userType = userType == 2 ? 2 : 1;
        String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        Workbook workbook = null;
        //读取Excel
        try {
            switch (suffix) {
                case "xls":
//                    POIFSFileSystem fs = new POIFSFileSystem(inputStream);
//                    workbook = new HSSFWorkbook(fs);
//                    workbook = new XSSFWorkbook(inputStream);
                    workbook = new HSSFWorkbook(inputStream);
                    break;
                case "xlsx":
                    workbook = new XSSFWorkbook(inputStream);
                    break;
                default:
                    break;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (workbook == null) {
            return ExcelResultDto.fail("导入文件不存在或类型不正确");
        }
        Sheet sheetAt = workbook.getSheetAt(0);
        //总行数
        int totalRow = sheetAt.getLastRowNum() + 1;
        if (totalRow <= DefaultNumberRows) {
            return ExcelResultDto.fail("导入的文件中没有内容");
        }
        //每行元素
        Row row = sheetAt.getRow(1);
        //获取总列数
        int totalCell = 0;
        if (userType == 1) {
            if (row.getPhysicalNumberOfCells() != 10) {
                return ExcelResultDto.fail("请使用正确的模板");
            }
            if (getCellValue(row.getCell(0)).indexOf("姓名") == -1
                    || getCellValue(row.getCell(1)).indexOf("手机号") == -1
                    || getCellValue(row.getCell(2)).indexOf("学号") == -1
                    || getCellValue(row.getCell(3)).indexOf("邮箱") == -1
                    || getCellValue(row.getCell(4)).indexOf("性别") == -1
                    || getCellValue(row.getCell(5)).indexOf("行政班/组织架构") == -1
                    || getCellValue(row.getCell(6)).indexOf("出生日期") == -1
                    || getCellValue(row.getCell(7)).indexOf("QQ") == -1
                    || getCellValue(row.getCell(8)).indexOf("描述") == -1
                    || getCellValue(row.getCell(9)).indexOf("密码") == -1) {
                return ExcelResultDto.fail("请使用正确的模板");
            }
        } else {
            if (row.getPhysicalNumberOfCells() != 12) {
                return ExcelResultDto.fail("请使用正确的模板");
            }
            if (getCellValue(row.getCell(0)).indexOf("姓名") == -1
                    || getCellValue(row.getCell(1)).indexOf("手机号") == -1
                    || getCellValue(row.getCell(2)).indexOf("工号") == -1
                    || getCellValue(row.getCell(3)).indexOf("是否建课") == -1
                    || getCellValue(row.getCell(4)).indexOf("邮箱") == -1
                    || getCellValue(row.getCell(5)).indexOf("性别") == -1
                    || getCellValue(row.getCell(6)).indexOf("院系/组织架构") == -1
                    || getCellValue(row.getCell(7)).indexOf("出生日期") == -1
                    || getCellValue(row.getCell(8)).indexOf("QQ") == -1
                    || getCellValue(row.getCell(9)).indexOf("描述") == -1
                    || getCellValue(row.getCell(10)).indexOf("职称") == -1
                    || getCellValue(row.getCell(11)).indexOf("密码") == -1) {
                return ExcelResultDto.fail("请使用正确的模板");
            }
        }
        List<ExcelUserDto> data = new ArrayList<>();
        List<ExcelErrorDto> error = new ArrayList<>();
        List<String> excelUserName = new ArrayList<>();
        List<String> excelUserNo = new ArrayList<>();
        List<String> excelEmail = new ArrayList<>();
        List<String> excelMobile = new ArrayList<>();
        int nullCount = 0;
        for (int i = DefaultNumberRows; i < totalRow; i++) {
            if(nullCount > 3){
                break;
            }
            row = sheetAt.getRow(i);
            if(row == null){
                nullCount++;
                continue;
            }
            nullCount = 0;
            ExcelUserDto info = new ExcelUserDto();
            info.setRows(i + 1);
            info.setDisplayName(getCellValue(row.getCell(0)));
            info.setMobile(getCellValue(row.getCell(1)));
            info.setUserNo(getCellValue(row.getCell(2)));

            if (userType == 1) {
                info.setEmail(getCellValue(row.getCell(3)));
                info.setSex(getCellValue(row.getCell(4)));
                info.setRedisName(getCellValue(row.getCell(5)));
                info.setBirthday(getCellValue(row.getCell(6)));
                info.setQq(getCellValue(row.getCell(7)));
                info.setRemark(getCellValue(row.getCell(8)));
                info.setPassword(getCellValue(row.getCell(9)));
                info.setDoCourse("否");
            } else {
                info.setEmail(getCellValue(row.getCell(4)));
                info.setDoCourse(getCellValue(row.getCell(3)));
                info.setSex(getCellValue(row.getCell(5)));
                info.setRedisName(getCellValue(row.getCell(6)));
                info.setBirthday(getCellValue(row.getCell(7)));
                info.setQq(getCellValue(row.getCell(8)));
                info.setRemark(getCellValue(row.getCell(9)));
                info.setPost(getCellValue(row.getCell(10)));
                info.setPassword(getCellValue(row.getCell(11)));
            }

            if (StringUtils.isEmpty(info.getDisplayName()) && StringUtils.isEmpty(info.getMobile()) && StringUtils.isEmpty(info.getEmail())) {
                continue;
            }
            data.add(info);
            if(data.size() > 3000){
                return ExcelResultDto.fail("一次最多导入3000条数据！");
            }
            int finalUserType = userType;
            //姓名
            if (StringUtils.isEmpty(info.getDisplayName())) {
                error.add(new ExcelErrorDto(i + 1, info.getDisplayName(), "该姓名不能为空"));
            } else if (info.getDisplayName().length() > 50) {
                error.add(new ExcelErrorDto(i + 1, info.getDisplayName(), "该姓名的长度不能大于50个字符"));
            }
            info.setUserName(info.getMobile());
            //手机号
            if (!StringUtils.isEmpty(info.getMobile())) {
//                if (info.getMobile().length() > 20) {
//                    error.add(new ExcelErrorDto(i + 1, info.getMobile(), "该手机号格式不正确"));
//                } else
                    if (excelMobile.stream().filter(q->q.equals(info.getMobile())).count() > 0) {
                    error.add(new ExcelErrorDto(i + 1, info.getMobile(), "该手机号在表格中存在"));
                }else if(userInfoDtos.stream().filter(q->StringUtils.isNotEmpty(q.getMobile()) && q.getMobile().equals(info.getMobile()) && q.getUserType() == finalUserType).count() > 0){
                    error.add(new ExcelErrorDto(i + 1, info.getMobile(), "该手机号在系统中已存在"));
                }
            }

            //用户名
            if (StringUtils.isEmpty(info.getUserName())) {
                error.add(new ExcelErrorDto(i + 1, info.getUserName(), "请填写手机号"));
            } else if (info.getUserName().length() > 50) {
                error.add(new ExcelErrorDto(i + 1, info.getUserName(),  "手机号的格式不正确"));
            } else if (CheckDataUtils.isContainChinese(info.getUserName())) {
                error.add(new ExcelErrorDto(i + 1, info.getUserName(), "手机号的格式不正确"));
            } else if (StringUtils.isNotEmpty(CheckDataUtils.validateLegalString(info.getUserName().replace("@", "")))) {
                error.add(new ExcelErrorDto(i + 1, info.getUserName(),  "手机号的格式不正确"));
            }
//            else if (excelUserName.contains(info.getUserName())) {
//                error.add(new ExcelErrorDto(i + 1, info.getUserName(),  "该手机号在表格中已存在"));
//            }
            else if(userInfoDtos.stream().filter(q->StringUtils.isNotEmpty(q.getUserName()) &&  q.getUserName().equals(info.getUserName()) && q.getUserType() == finalUserType).count() > 0){
                error.add(new ExcelErrorDto(i + 1, info.getUserName(), "该手机号作为用户名在系统中已存在"));
            }
            excelUserName.add(info.getUserName());

            //工号/学号
            if (StringUtils.isEmpty(info.getUserNo())) {
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "不能为空"));
            } else if (info.getUserNo().length() > 50) {
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "的长度不能大于50个字符"));
            } else if (CheckDataUtils.isContainChinese(info.getUserNo())) {
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "中不能包含中文"));
            } else if (StringUtils.isNotEmpty(CheckDataUtils.validateLegalString(info.getUserNo()))) {
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "中不能包含非法字符"));
            } else if (excelUserNo.contains(info.getUserNo())){
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "在该表格中已存在"));
            }else if(userInfoDtos.stream().filter(q->StringUtils.isNotEmpty(q.getEmployeeNumber()) && q.getEmployeeNumber().equals(info.getUserNo())).count() > 0){
                error.add(new ExcelErrorDto(i + 1, info.getUserNo(), (userType == 1 ? "该学号" : "该工号") + "在系统中已存在"));
            }
            excelUserNo.add(info.getUserNo());


            if (StringUtils.isNotEmpty(info.getMobile())) {
                excelMobile.add(info.getMobile());
            }
            //邮箱
            if (!StringUtils.isEmpty(info.getEmail())) {
                if (info.getEmail().length() > 50 || !info.getEmail().contains("@")) {
                    error.add(new ExcelErrorDto(i + 1, info.getEmail(), "该邮箱格式不正确"));
                } else if (excelEmail.equals(info.getEmail())) {
                    error.add(new ExcelErrorDto(i + 1, info.getEmail(), "该邮箱在表格中存在"));
                }else if(userInfoDtos.stream().filter(q->StringUtils.isNotEmpty(q.getEmail()) && q.getEmail().equals(info.getEmail())  && q.getUserType() == finalUserType).count() > 0){
                    error.add(new ExcelErrorDto(i + 1, info.getEmail(), "该邮箱在系统中已存在"));
                }
            }
            if (StringUtils.isNotEmpty(info.getEmail())) {
                excelEmail.add(info.getEmail());
            }
            //性别
            if (StringUtils.isNotEmpty(info.getSex())) {
                if (!info.getSex().equals("男") && !info.getSex().equals("女")) {
                    error.add(new ExcelErrorDto(i + 1, info.getSex(), "该性别格式不正确（只能为“男”或“女”）"));
                }
            }
            //行政班/组织架构
            if (StringUtils.isNotEmpty(info.getRedisName())) {
                if (!redisDtoList.isEmpty()) {
                    if (schoolType == 4) { //企业
                        String[] arr = info.getRedisName().split("/");
                        String d = getResId(redisDtoList, 1, arr, arr.length - 1, new ArrayList<>());
                        if (d.contains("不存在")) {
                            error.add(new ExcelErrorDto(i + 1, info.getRedisName(), (userType == 1 ? "该行政班/组织架构" : "该院系/组织架构") + "在系统中不存在"));
                        } else {
                            info.setRedisId(d);
                        }
                    } else {
                        RedisDto classInfo = redisDtoList.stream().filter(q -> q.getName().equals(info.getRedisName())).findAny().orElse(null);
                        if (classInfo != null) {
                            if (userType == 1) {
                                info.setRedisId(classInfo.getId());
                                info.setSubId(classInfo.getSubId());
                                info.setMajorId(classInfo.getMajorId());
                                info.setGradeId(classInfo.getGradeId());
                            } else {
                                info.setRedisId(classInfo.getId());
                            }
                        } else {
                            error.add(new ExcelErrorDto(i + 1, info.getRedisName(), (userType == 1 ? "该行政班/组织架构" : "该院系/组织架构") + "在系统中不存在"));
                        }
                    }
                } else {
                    error.add(new ExcelErrorDto(i + 1, info.getRedisName(), (userType == 1 ? "该行政班/组织架构" : "该院系/组织架构") + "在系统中不存在"));
                }
            }
            // 出生日期
            // QQ
            if (StringUtils.isNotEmpty(info.getQq()) && info.getQq().length() > 50) {
                error.add(new ExcelErrorDto(i + 1, info.getMobile(), "该QQ的长度不能大于50个字符"));
            }
            // 描述
            //建课权限
            if (userType == 2) {
                if (StringUtils.isEmpty(info.getDoCourse())) {
                    error.add(new ExcelErrorDto(i, info.getDoCourse(), "该建课权限不能为空"));
                } else if (!info.getDoCourse().equals("是") && !info.getDoCourse().equals("否")) {
                    error.add(new ExcelErrorDto(i + 1, info.getDoCourse(), "该建课权限格式不正确（只能为“是”或“否”）"));
                }
            }
            // 职称
            //密码
            if (StringUtils.isNotBlank(info.getPassword()) && !PasswordUtil.excelStrongCheck(info.getPassword())) {
                error.add(new ExcelErrorDto(i + 1, info.getPassword(), "改密码不符合填写须知中的规则"));
            }
        }
        if (error.isEmpty()) {
            return ExcelResultDto.success(data);
        } else {
            return ExcelResultDto.error(error);
        }
    }

    /**
     * 获取组织架构
     * `
     *
     * @param redisDtoList
     * @param i
     * @param arr
     * @param num
     * @param parentIds`
     * @return java.lang.String
     * @Description
     * @Author xzli
     * @CreateDate 2021/6/8 10:01
     */
    public static String getResId(List<RedisDto> redisDtoList, int i, String[] arr, int num, List<String> parentIds) {
        List<RedisDto> fistList = new ArrayList<>();
        if (parentIds.size() > 0) {
            fistList = redisDtoList.stream().filter(q -> q.getName().equals(arr[num]) && parentIds.contains(q.getId())).collect(Collectors.toList());
        } else {
            fistList = redisDtoList.stream().filter(q -> q.getName().equals(arr[num])).collect(Collectors.toList());
        }
        if (fistList.size() == 1) {
            return fistList.get(0).getNodePath();
        } else if (fistList.size() > 1 && arr.length - i > 0) {
            List<String> thisParentIds = fistList.stream().filter(q -> StringUtils.isNotEmpty(q.getParentId())).map(RedisDto::getParentId).collect(Collectors.toList());
            if (thisParentIds.size() == 0) {
                return "不存在";
            }
            return getResId(redisDtoList, i + 1, arr, num - 1, thisParentIds);
        } else {
            return "不存在";
        }
    }

    /**
     * 获取excel某个格的内容
     *
     * @param cell
     * @return java.lang.String
     * @Description
     * @Author xzli
     * @CreateDate 2021/6/8 10:01
     */
    public static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return null;
        }
        // 以下是判断数据的类型
        switch (cell.getCellTypeEnum()) {
            case NUMERIC: // 数字
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = sdf.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                } else {
                    DataFormatter dataFormatter = new DataFormatter();
                    cellValue = dataFormatter.formatCellValue(cell);
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                cellValue = cell.getCellFormula() + "";
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                //cellValue = "非法字符";
                cellValue = "";
                break;
            default:
                //cellValue = "未知类型";
                cellValue = "";
                break;
        }
        return cellValue.trim();
    }

    public static void getHSSFWorkbook(HttpServletResponse response, String sheetName, String[] title, String[][] values) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        //style.setAlignment(new HSSFCellStyle()); // 创建一个居中格式

        //声明列对象
        HSSFCell cell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        for (int i = 0; i < values.length; i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < values[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                row.createCell(j).setCellValue(values[i][j]);
            }
        }

        //响应到客户端
        try {
            ExcelUtils.setResponseHeader(response, sheetName + ".xls");
            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // return wb;

    }

    public static void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            try {
                fileName = new String(fileName.getBytes(), "ISO8859-1");
            } catch (UnsupportedEncodingException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            response.setContentType("application/octet-stream;charset=ISO8859-1");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
