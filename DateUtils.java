package com.mukun.user.config.utils;

import org.springframework.stereotype.Component;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;

@Component
public class DateUtils {

    /**
     * 某天零时零分零秒时间
     * @param time 时间
     **/
    public static Date startDate(Date time) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(time);
        calendar.set(Calendar.SECOND, 0);
        calendar.set(Calendar.MINUTE, 0);
        calendar.set(Calendar.HOUR_OF_DAY, 0);
        calendar.set(Calendar.MILLISECOND, 0);
        return calendar.getTime();
    }

    /**
     * 某天23时59分59秒
     * @param time 时间
     **/
    public static Date endDate(Date time) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(time);
        calendar.set(Calendar.SECOND, 59);
        calendar.set(Calendar.MINUTE, 59);
        calendar.set(Calendar.HOUR_OF_DAY, 23);
        calendar.set(Calendar.MILLISECOND, 0);
        return calendar.getTime();
    }

    /**
     * String转Date
     * @param dateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021/11/23 13:08
     * @Return
     */
    public static Date stringToDateTime(String dateTime){
        SimpleDateFormat ft = new SimpleDateFormat("yyyy-MM-dd");
        Date date = new Date();
        try {
            date = ft.parse(dateTime);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return date;
    }

    /**
     * Date转String(yyyy-MM-dd HH:mm:ss)
     * @param dateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021/11/23 13:10
     * @Return
     */
    public static String dateToStringTime (Date dateTime){
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateString = formatter.format(dateTime);
        return dateString;
    }

    /**
     * String转LocalDateTime
     * @param dateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021/11/23 16:01
     * @Return
     */
    public static LocalDateTime stringToLocalDateTime(String dateTime){
        DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        LocalDateTime localDateTime = LocalDateTime.parse(dateTime, df);
        return localDateTime;
    }

    /**
     * LocalDateTime转String
     * @param dateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021/11/23 16:06
     * @Return
     */
    public static String localDateTimeToString  (LocalDateTime dateTime){
        DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String localTimeString = df.format(dateTime);
        return localTimeString;
    }

    /**
     * Date转LocalDateTime
     * @param date Date
     * @Author xhzhang
     * @Description
     * @CreateDate 2021-11-23 19:23:48
     * @return LocalDateTime
     */
    public static LocalDateTime dateToLocalDateTime(Date date) {
        try {
            Instant instant = date.toInstant();
            ZoneId zoneId = ZoneId.systemDefault();
            return instant.atZone(zoneId).toLocalDateTime();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * LocalDateTime转Date
     * @param localDateTime LocalDateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021-11-23 19:23:57
     * @return Date
     */
    public static Date localDateTimeToDate(LocalDateTime localDateTime) {
        try {
            ZoneId zoneId = ZoneId.systemDefault();
            ZonedDateTime zdt = localDateTime.atZone(zoneId);
            return Date.from(zdt.toInstant());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * LocalDateTime转毫秒时间戳
     * @param localDateTime LocalDateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2021-11-23 19:40:12
     * @return 时间戳
     */
    public static Long localDateTimeToTimestamp(LocalDateTime localDateTime) {
        try {
            ZoneId zoneId = ZoneId.systemDefault();
            Instant instant = localDateTime.atZone(zoneId).toInstant();
            return instant.toEpochMilli();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 时间戳转LocalDateTime
     * @param timestamp 时间戳
     * @Author xhzhang
     * @Description
     * @CreateDate 2021-11-23 19:41:12
     * @return LocalDateTime
     */
    public static LocalDateTime timestampToLocalDateTime(long timestamp) {
        try {
            Instant instant = Instant.ofEpochMilli(timestamp);
            ZoneId zone = ZoneId.systemDefault();
            return LocalDateTime.ofInstant(instant, zone);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * Date转String(yyyyMMddHHmmss)
     * @param dateTime
     * @Author xhzhang
     * @Description
     * @CreateDate 2022-2-17 09:57:54
     * @Return
     */
    public static String dateToStringTimeTwo (Date dateTime){
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddHHmmss");
        String dateString = formatter.format(dateTime);
        return dateString;
    }

}
