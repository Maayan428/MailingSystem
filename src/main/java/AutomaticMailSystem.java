import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.*;
import javax.mail.internet.*;
import java.io.*;
import java.util.*;

public class AutomaticMailSystem {


    static class EmailData {
        String email;
        String subjectName;
        String companyName;

        EmailData(String email, String subjectName, String companyName) {
            this.email = email;
            this.subjectName = subjectName;
            this.companyName = companyName;
        }
    }

    public static void main(String[] args) {
        String excelFilePath = "src/main/resources/email_list.xlsx"; // Update the path
        String mailBodyFilePath = "src/main/resources/mail_body.txt"; // Path to the .txt file
        String username = "pmashabim@sn.matnasim.co.il";
        String appPassword = "Shlo1373";

        List<File> attachments = List.of(
                new File("/Users/maayanturgeman/IdeaProjects/Mailing_system/src/main/resources/attachments/ הכשרת מטפלים.pdf"),
                new File("/Users/maayanturgeman/IdeaProjects/Mailing_system/src/main/resources/attachments/ פיזיקה.pdf"),
                new File("/Users/maayanturgeman/IdeaProjects/Mailing_system/src/main/resources/attachments/אישור בעלות חן מתנס.pdf"),
                new File("/Users/maayanturgeman/IdeaProjects/Mailing_system/src/main/resources/attachments/אישור 46 -580370781.pdf"),
                new File("/Users/maayanturgeman/IdeaProjects/Mailing_system/src/main/resources/attachments/מרכז טיפולי מוטלידסיפלינארי.pdf")
        );

        try {
            List<EmailData> emailDataList = readExcelFile(excelFilePath);
            String mailBodyTemplate = readMailBodyTemplate(mailBodyFilePath);
            sendEmails(emailDataList, mailBodyTemplate, username, appPassword, attachments, excelFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<EmailData> readExcelFile(String filePath) throws Exception {
        List<EmailData> emailDataList = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                String email = row.getCell(0).getStringCellValue();
                String subjectName = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : "";
                String companyName = row.getCell(2).getStringCellValue();

                emailDataList.add(new EmailData(email, subjectName, companyName));
            }
        }
        return emailDataList;
    }

    private static String readMailBodyTemplate(String filePath) throws IOException {
        StringBuilder content = new StringBuilder();
        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            String line;
            while ((line = br.readLine()) != null) {
                content.append(line).append("\n");
            }
        }
        return content.toString();
    }

    private static void sendEmails(List<EmailData> emailDataList, String mailBodyTemplate, String username, String appPassword, List<File> attachments, String excelFilePath) {
        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.office365.com");
        props.put("mail.smtp.port", "587");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");

        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, appPassword);
            }
        });

        for (EmailData emailData : emailDataList) {
            try {
                String personalizedBody = mailBodyTemplate.replace("[SUBJECT_NAME]", emailData.subjectName)
                        .replace("[COMPANY_NAME]", emailData.companyName);

                Message message = new MimeMessage(session);
                message.setFrom(new InternetAddress(username));
                message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(emailData.email));
                message.setSubject("נושא לשליחה");

                MimeBodyPart messageBodyPart = new MimeBodyPart();
                messageBodyPart.setText(personalizedBody, "utf-8");

                Multipart multipart = new MimeMultipart();
                multipart.addBodyPart(messageBodyPart);

                for (File attachment : attachments) {
                    MimeBodyPart attachmentPart = new MimeBodyPart();
                    attachmentPart.attachFile(attachment);
                    multipart.addBodyPart(attachmentPart);
                }

                message.setContent(multipart);

                Transport.send(message);
                updateExcelStatus(emailData, "Sent", excelFilePath);
                System.out.println("Email sent to: " + emailData.email);

            } catch (Exception e) {
                System.err.println("Failed to send email to: " + emailData.email);
                e.printStackTrace();
                updateExcelStatus(emailData, "Failed", excelFilePath);
            }
        }
    }

    private static void updateExcelStatus(EmailData emailData, String status, String excelFilePath) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                String email = row.getCell(0).getStringCellValue();
                if (email.equals(emailData.email)) {
                    Cell statusCell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    statusCell.setCellValue(status);
                    break;
                }
            }
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

