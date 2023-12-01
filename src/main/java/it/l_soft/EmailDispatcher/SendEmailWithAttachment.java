package it.l_soft.EmailDispatcher;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.log4j.Logger;

public class SendEmailWithAttachment {
	private static final Logger logger = Logger.getLogger(SendEmailWithAttachment.class);


//	private static String toNamecase(String s)
//	{
//		String sep = "";
//		String retValue = "";
//		s = s.toLowerCase();
//		String[] items = s.split(" ");
//		for(String item : items)
//		{
//			if (item.trim().length() == 0)
//				continue;
//			retValue = retValue + sep + item.substring(0, 1).toUpperCase().concat(item.substring(1));
//			sep = " ";
//		}
//		return(retValue);
//	}
	
	public SendEmailWithAttachment(String[] args) {
		int excelHeaderRow = -1; // la riga su cui si trova il column header per scegliere la colonna email 
		int startFrom = -1; // la prima riga da considerare su excel per valutare la spedizione
		int howMany = -1;
		int timeout = -1;
		int countSent = 0;
		ApplicationProperties ap = ApplicationProperties.getInstance();
		
		try {
			if (args.length < 5)
			{
				logger.debug("Need a record to start from in the excel file");
				logger.debug("Usage: SendEmailWithAttachment header_row record_to_start_from how_many timeout checkBoxIdx");
				System.exit(-1);
			}
			excelHeaderRow = Integer.parseInt(args[0]);
			startFrom = Integer.parseInt(args[1]);
			howMany = Integer.parseInt(args[2]);
			timeout = Integer.parseInt(args[3]);
			ap.setEmailSendCheckbox(Integer.parseInt(args[4]));
			if (args.length == 6)

				ap.setEmailCustomersDismissed(Integer.parseInt(args[5]));
		}
		catch(Exception e)
		{
			logger.debug("Exception " + e.getMessage() + " converting '" + args[0] + "' to int");
		}
		finally
		{
			if ((startFrom == -1) || (howMany == -1) || (timeout == -1))
			{
				logger.debug("Usage: SendEmailWithAttachment record_to_start_from how_many timeout");
				System.exit(-1);
			}
		}
		ReadFromExcel excel = null;
		try {
			excel = new ReadFromExcel(excelHeaderRow, startFrom, ap);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Properties props = new Properties();

		logger.debug("Connecting to '" + ap.getMailServerHost() + "' on port " + ap.getMailServerPort());
		props.put("mail.smtp.host", ap.getMailServerHost());
		props.put("mail.smtp.ssl.trust", ap.getMailServerHost());
		props.put("mail.smtp.auth", "true");
//		props.put("mail.smtp.auth", "false");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.ssl.protocols", "TLSv1.2");
		props.put("mail.smtp.port", ap.getMailServerPort());

		props.put("mail.smtp.user", ap.getMailServerUsername());
		props.put("mail.smtp.password", ap.getMailServerPassword());

		BufferedReader br = null;
		StringBuilder mailBody = new StringBuilder();
		try {
			br = new BufferedReader(new FileReader(ap.getMailBodyPath()));
			String line = br.readLine();

			while (line != null) {
				mailBody.append(line);
				mailBody.append("\n");
				line = br.readLine();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// Get the Session object.
		Session session = Session.getInstance(props,
				new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(ap.getMailServerUsername(), ap.getMailServerPassword());
			}
		});
//		Session session = Session.getDefaultInstance(props);

		int count = 0;
		try {
			while(excel.getNextRow() != null)
			{
				if ((excel.getField(ap.getEmailSendCheckbox()).toUpperCase().compareTo("X") == 0) || 
					(excel.getField(ap.getEmailCustomersDismissed()).toUpperCase().compareTo("X") == 0) || 
					(excel.getEmail().compareTo("") == 0) ||
					((ap.getEmailSentCheckbox() >= 0) && (excel.getField(ap.getEmailSentCheckbox()).compareTo("*") == 0)))
					continue;
				if (count++ == howMany)
				{
					logger.debug("Sent " + howMany + " emails. Reached the row " + 
									   excel.rowIdx + " on the file. Quitting the job");
					break;
				}
				
				if (countSent++ == ap.getStopEvery())
				{
					try
					{
						Thread.sleep(ap.getStopFor());
					}
					catch(Exception e)
					{
						;
					}
					countSent = 1;
				}
				
				String cliente = excel.getField(ap.getColHeaders()[Constants.COGNOME]);
				if (cliente.length() > 15)
				{
					cliente = cliente.substring(0, 15);
				}
				else if (cliente.length() == 0)
				{
					cliente = excel.getEmail();
				}
				
				logger.debug("Row " + excel.rowIdx + " - client " + cliente );
				// Create a default MimeMessage object.
				Message message = new MimeMessage(session);

				// Set From: header field of the header.
				message.setFrom(new InternetAddress(ap.getEmailSentFrom().compareTo("") == 0 ? 
											ap.getMailServerUsername() : ap.getEmailSentFrom()));

				if ((excel.getEmail() == null) || 
					(excel.getEmail().trim().compareTo("") == 0) ||
					(excel.getEmail().trim().compareTo("0") == 0))
				{
					logger.debug(". Empty email field - *** NOT SENDING ***");
					continue;
				}

				// Set To: header field of the header.
				message.setRecipients(Message.RecipientType.TO,
							InternetAddress.parse(excel.getEmail()));
//							InternetAddress.parse("osvaldo.lucchini@gmail.com"));
				
				if (excel.getEmailCCValue() != null)
				{
					message.setRecipients(Message.RecipientType.CC,
							InternetAddress.parse(excel.getEmailCCValue()));
//							InternetAddress.parse("osvaldo.lucchini@wedi.it"));
				}
				
				if (excel.getEmailBCCValue() != null)
				{
					message.setRecipients(Message.RecipientType.BCC,
								InternetAddress.parse(excel.getEmailBCCValue()));
				}

				// Set Subject: header field
				message.setSubject(ap.getMailSubject());

				// Create a multipar message and the message part
				Multipart multipart = new MimeMultipart();
				BodyPart messageBodyPart = new MimeBodyPart();
				DataSource source;
				try
				{
					// Now set the actual message
					String body = mailBody.toString();
					for(String colName: ap.getColHeaders())
					{
						body = body.replaceAll("\\$" + colName + "\\$", excel.getField(colName));
					}

					messageBodyPart.setContent(body, "text/html");
					multipart.addBodyPart(messageBodyPart);

//					// Part two is the picture
					MimeBodyPart imagePart = new MimeBodyPart();
					imagePart = new MimeBodyPart();
					imagePart.attachFile(ap.getCommonDocsFolder() + "wedi.png");
					imagePart.setContentID("<wediLogo>");
					imagePart.setDisposition(MimeBodyPart.INLINE);
					multipart.addBodyPart(imagePart);

//					// Another picture
//					imagePart = new MimeBodyPart();
//					imagePart.attachFile("." + File.separator + "docs" + File.separator + "aew21" + File.separator + "aawlogo.png");
//					imagePart.setContentID("<aawlogo>");
//					imagePart.setDisposition(MimeBodyPart.INLINE);
//					multipart.addBodyPart(imagePart);
//
					// signature needed?
					if (ap.getSignatureFilePath() != null)
					{
						imagePart = new MimeBodyPart();
						imagePart.attachFile(ap.getCommonDocsFolder() + ap.getSignatureFilePath());
						imagePart.setContentID("<signature>");
						imagePart.setDisposition(MimeBodyPart.INLINE);
						multipart.addBodyPart(imagePart);
					}
					
					// attachments needed?
					int y = 0;
					String[] attachments = excel.getFileAttachValue().split(";");					
					while((y < attachments.length) && attachments[y].compareTo("") != 0)
					{
						messageBodyPart = new MimeBodyPart();
						source = new FileDataSource(ap.getExcelEmailAttachFilesFolder() + attachments[y]);
						messageBodyPart.setDataHandler(new DataHandler(source));
						if (attachments[y].lastIndexOf(File.separator) >= 0)
							messageBodyPart.setFileName(attachments[y].substring(attachments[y].lastIndexOf(File.separator)));
						else
							messageBodyPart.setFileName(attachments[y]);
						multipart.addBodyPart(messageBodyPart);
						y++;
					}

					// Send the complete message parts
					message.setContent(multipart);

					logger.debug(" sending to:");
					for(y = 0; y < message.getAllRecipients().length; y++)
						logger.debug(message.getAllRecipients()[y]);
					Transport.send(message);
					logger.debug(" - sent successfully....");
					excel.setSentFlag(ap.getEmailSentCheckbox(), "*");
					Thread.sleep(timeout);
				}
				catch(Exception e)
				{
					logger.debug(" Mail to " + excel.getEmail() + " not sent. Exception " + e.getMessage());
					e.printStackTrace();
				}
			}
		} catch (MessagingException e) {
			throw new RuntimeException(e);
		}
		finally
		{
			try
			{
				excel.writeChanges();
				logger.debug("Changes saved");
			}
			catch(Exception e)
			{
				logger.debug("Changes not written on the source file. Exception " + e.getMessage());
			}
		}
	}
}
