package com.mukun.user.config.utils;

import org.springframework.context.annotation.Configuration;

import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.Properties;

@Configuration
public class SendEmailUtil {

    public static void sendEmail(String myEmailAccount, String myEmailPassword, String email, String body) throws UnsupportedEncodingException, MessagingException {
        try {

            Thread thread = new Thread(() -> {
                try {
                    sendEmailMain(myEmailAccount, myEmailPassword, email, body);
                } catch (UnsupportedEncodingException e) {
                    e.printStackTrace();
                } catch (MessagingException e) {
                    e.printStackTrace();
                }
            });
            thread.start();

        } catch (Exception e) {

        }
    }

    /**
     * @param myEmailAccount
     * @param myEmailPassword
     * @param email
     * @param body
     * @Description
     * @Author xzli
     * @CreateDate 2021/5/24 18:04
     */
    public static void sendEmailMain(String myEmailAccount, String myEmailPassword, String email, String body) throws UnsupportedEncodingException, MessagingException {
        Properties props = new Properties();
        props.setProperty("mail.transport.protocol", "smtp");
//        props.setProperty("mail.smtp.host", "smtp.qq.com");
        props.setProperty("mail.smtp.host", "smtp.exmail.qq.com");
        props.setProperty("mail.smtp.auth", "true");
        final String smtpPort = "465";
        props.setProperty("mail.smtp.port", smtpPort);
        props.setProperty("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        props.setProperty("mail.smtp.socketFactory.fallback", "false");
        props.setProperty("mail.smtp.socketFactory.port", smtpPort);
        Session session = Session.getDefaultInstance(props);
        session.setDebug(true);                                 // 设置为debug模式, 可以查看详细的发送 log
        MimeMessage message = createMimeMessage(session, myEmailAccount, email, body);
        Transport transport = session.getTransport();
        transport.connect(myEmailAccount, myEmailPassword);
        transport.sendMessage(message, message.getAllRecipients());
        transport.close();
    }

    public static MimeMessage createMimeMessage(Session session, String sendMail, String receiveMail, String text) throws UnsupportedEncodingException, MessagingException {
        MimeMessage message = new MimeMessage(session);
        message.setFrom(new InternetAddress(sendMail, "发送人名称", "UTF-8"));
        message.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(receiveMail, "", "UTF-8"));
        message.setSubject("功能提示语", "UTF-8");
        message.setContent(text, "text/html;charset=UTF-8");
        message.setSentDate(new Date());
        message.saveChanges();
        return message;
    }
}
