package com.mukun.user.config.utils;

import com.baomidou.mybatisplus.core.toolkit.StringUtils;
import com.mukun.user.api.model.BusinessException;
import org.apache.commons.codec.digest.DigestUtils;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * password utl
 */
public class PasswordUtil {

    private static final Set<Character> specialCharacter
            = new HashSet<>(Arrays.asList('!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '-', '='));

    private static final String PASSWORD_STRONG = "密码由大写字母、小写字母、数字、特殊字符中任意三种组成，长度为 8-32 位 ，特殊字符为：!@#$%^&*()_+-=";

    private PasswordUtil() {
    }

    /**
     * 生成密文密码
     *
     * @param password 明文密码
     */
    public static String encrypt(String password) {
        return DigestUtils.md5Hex(password);
    }

    /**
     * 密码强度验证
     *
     * @param password 明文密码
     */
    public static void strongCheck(String password) {
        if (StringUtils.isBlank(password)) {
            throw new BusinessException(PASSWORD_STRONG);
        }

        if ((password.length() < 8) || (password.length() > 32)) {
            throw new BusinessException(PASSWORD_STRONG);
        }

        boolean hasUpperCase = false;
        boolean hasLowerCase = false;
        boolean hasNumber = false;
        boolean hasSpecialCharacter = false;

        for (int i = 0; i < password.length(); i++) {
            char c = password.charAt(i);
            if (Character.isUpperCase(c)) {
                hasUpperCase = true;
            } else if (Character.isLowerCase(c)) {
                hasLowerCase = true;
            } else if (Character.isDigit(c)) {
                hasNumber = true;
            } else if (specialCharacter.contains(c)) {
                hasSpecialCharacter = true;
            } else {
                throw new BusinessException(PASSWORD_STRONG);
            }
        }

        int level = 0;

        if (hasUpperCase) {
            level += 1;
        }

        if (hasLowerCase) {
            level += 1;
        }

        if (hasNumber) {
            level += 1;
        }

        if (hasSpecialCharacter) {
            level += 1;
        }

        if (level < 3) {
            throw new BusinessException(PASSWORD_STRONG);
        }
    }

    /**
     * Excel导入密码强度验证
     * @param password
     * @Author xhzhang
     * @Description
     * @CreateDate 2021/12/13 15:11
     * @Return
     */
    public static boolean excelStrongCheck(String password) {
        if (StringUtils.isBlank(password)) {
            return false;
        }
        if ((password.length() < 8) || (password.length() > 32)) {
            return false;
        }
        boolean hasUpperCase = false;
        boolean hasLowerCase = false;
        boolean hasNumber = false;
        boolean hasSpecialCharacter = false;
        for (int i = 0; i < password.length(); i++) {
            char c = password.charAt(i);
            if (Character.isUpperCase(c)) {
                hasUpperCase = true;
            } else if (Character.isLowerCase(c)) {
                hasLowerCase = true;
            } else if (Character.isDigit(c)) {
                hasNumber = true;
            } else if (specialCharacter.contains(c)) {
                hasSpecialCharacter = true;
            } else {
                return false;
            }
        }
        int level = 0;
        if (hasUpperCase) {
            level += 1;
        }
        if (hasLowerCase) {
            level += 1;
        }
        if (hasNumber) {
            level += 1;
        }
        if (hasSpecialCharacter) {
            level += 1;
        }
        if (level < 3) {
            return false;
        }
        return true;
    }

}
