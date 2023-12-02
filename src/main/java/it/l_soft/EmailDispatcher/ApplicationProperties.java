package it.l_soft.EmailDispatcher;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.log4j.Logger;

public class ApplicationProperties {
	private int emailSendCheckbox = 0; 
	private int emailSentCheckbox = 1; 
	private int emailCustomersDismissed = 3; 

	private String emailSentFrom = ""; 
	private String excelEmailTOFieldName = "mailTo";
	private String excelEmailCCFieldName = "mailCC";  	
	
	private String excelEmailBCCFieldName = "commerciale@wedi.it";
	private String commonDocsFolder = "." + File.separator + "docs" + File.separator + "commonDocs" + File.separator;
	private String excelEmailAttachFilesFolder = "";
	private String excelEmailAttachFilesExt = "";
	private String excelEmailAttachFieldName = "";
	
	private String signatureFilePath = "";
	private String sheetName = "";
	private String excelFilePath = commonDocsFolder + "listaClienti.xls";
	private String mailBodyPath = excelEmailAttachFilesFolder + "testomail.html";
	private String mailSubject = "";
	
	private String mailServerUsername = ""; 
	private String mailServerHost = "";  
	private String mailServerPort = "";
	private String mailServerPassword = ""; 
	
	private int stopEvery = 50;
	private int stopFor = 5000;
	
	private boolean useAuth = false;
	
	private String[] colHeaders;
	
	private static ApplicationProperties instance = null;
	
	final Logger log = Logger.getLogger(this.getClass());
	
	public static ApplicationProperties getInstance()
	{
		if (instance == null)
		{
			instance = new ApplicationProperties();
		}
		return(instance);
	}
	
	private ApplicationProperties()
	{		
		log.trace("ApplicationProperties start");
		Properties properties = new Properties();
    	try 
    	{
    		log.debug("opening 'package.properties'");
        	InputStream in = ApplicationProperties.class.getResourceAsStream("/package.properties");
        	if (in == null)
        	{
        		log.error("resource path not found");
        		return;
        	}
        	properties.load(in);
	    	in.close();
		}
    	catch(IOException e) 
    	{
			log.warn("Exception " + e.getMessage(), e);
    		return;
		}
    	emailSendCheckbox = Integer.parseInt(properties.getProperty("emailSendCheckbox")); 
    	emailSentCheckbox = Integer.parseInt(properties.getProperty("emailSentCheckbox"));
    	emailCustomersDismissed = Integer.parseInt(properties.getProperty("emailCustomersDismissed")); 

    	emailSentFrom = properties.getProperty("emailSentFrom");
    	excelEmailTOFieldName = properties.getProperty("excelEmailTOFieldName");
    	excelEmailCCFieldName = properties.getProperty("excelEmailCCFieldName");

    	excelEmailBCCFieldName = properties.getProperty("excelEmailBCCFieldName");
    	commonDocsFolder = properties.getProperty("commonDocsFolder");
    	excelEmailAttachFilesFolder = properties.getProperty("excelEmailAttachFilesFolder");
    	excelEmailAttachFilesExt = properties.getProperty("excelEmailAttachFilesExt");
    	excelEmailAttachFieldName = properties.getProperty("excelEmailAttachFieldName");

    	signatureFilePath = properties.getProperty("signatureFilePath");
    	sheetName = properties.getProperty("sheetName");
    	excelFilePath = commonDocsFolder + File.separator + properties.getProperty("excelFilePath");
    	mailBodyPath = excelEmailAttachFilesFolder + properties.getProperty("mailBodyPath");
    	mailSubject = properties.getProperty("mailSubject");

    	stopEvery = Integer.parseInt(properties.getProperty("stopEvery"));
    	stopFor = Integer.parseInt(properties.getProperty("stopFor"));
    	colHeaders = properties.getProperty("colHeaders").split(", ");
    	
    	useAuth = Boolean.parseBoolean(properties.getProperty("useAuth"));
    	
    	String envConf = System.getProperty("envConf");
    	try 
    	{
    		properties = new Properties();
    		String siteProps = "/site." + (envConf == null ? "dev" : envConf) + ".properties";
    		log.debug("using '" + siteProps + "' at '" + ApplicationProperties.class.getResource(siteProps).getPath() + "'");
        	InputStream in = ApplicationProperties.class.getResourceAsStream(siteProps);        	
			properties.load(in);
	    	in.close();
		}
    	catch(IOException e) 
    	{
			log.error("Exception " + e.getMessage(), e);
    		return;
		}
    	mailServerHost = properties.getProperty("mailServerHost");
    	mailServerPort = properties.getProperty("mailServerPort");

    	properties = new Properties();
    	try 
    	{
    		log.debug("reading credentials");
        	InputStream in = ApplicationProperties.class.getResourceAsStream("/credentials");
        	if (in == null)
        	{
        		log.error("resource path not found");
        		return;
        	}
        	properties.load(in);
	    	in.close();
		}
    	catch(IOException e) 
    	{
			log.warn("Exception " + e.getMessage(), e);
    		return;
		}
    	mailServerUsername = properties.getProperty("mailServerUsername");
    	mailServerPassword = properties.getProperty("mailServerPassword");
	}

	public int getEmailSendCheckbox() {
		return emailSendCheckbox;
	}

	public void setEmailSendCheckbox(int emailSendCheckbox) {
		this.emailSendCheckbox = emailSendCheckbox;
	}

	public int getEmailSentCheckbox() {
		return emailSentCheckbox;
	}

	public void setEmailSentCheckbox(int emailSentCheckbox) {
		this.emailSentCheckbox = emailSentCheckbox;
	}

	public int getEmailCustomersDismissed() {
		return emailCustomersDismissed;
	}

	public void setEmailCustomersDismissed(int emailCustomersDismissed) {
		this.emailCustomersDismissed = emailCustomersDismissed;
	}

	public String getEmailSentFrom() {
		return emailSentFrom;
	}

	public void setEmailSentFrom(String emailSentFrom) {
		this.emailSentFrom = emailSentFrom;
	}

	public String getExcelEmailTOFieldName() {
		return excelEmailTOFieldName;
	}

	public void setExcelEmailTOFieldName(String excelEmailTOFieldName) {
		this.excelEmailTOFieldName = excelEmailTOFieldName;
	}

	public String getExcelEmailCCFieldName() {
		return excelEmailCCFieldName;
	}

	public void setExcelEmailCCFieldName(String excelEmailCCFieldName) {
		this.excelEmailCCFieldName = excelEmailCCFieldName;
	}

	public String getExcelEmailBCCFieldName() {
		return excelEmailBCCFieldName;
	}

	public void setExcelEmailBCCFieldName(String excelEmailBCCFieldName) {
		this.excelEmailBCCFieldName = excelEmailBCCFieldName;
	}

	public String getCommonDocsFolder() {
		return commonDocsFolder;
	}

	public void setCommonDocsFolder(String commonDocsFolder) {
		this.commonDocsFolder = commonDocsFolder;
	}

	public String getExcelEmailAttachFilesFolder() {
		return excelEmailAttachFilesFolder;
	}

	public void setExcelEmailAttachFilesFolder(String excelEmailAttachFilesFolder) {
		this.excelEmailAttachFilesFolder = excelEmailAttachFilesFolder;
	}

	public String getExcelEmailAttachFilesExt() {
		return excelEmailAttachFilesExt;
	}

	public void setExcelEmailAttachFilesExt(String excelEmailAttachFilesExt) {
		this.excelEmailAttachFilesExt = excelEmailAttachFilesExt;
	}

	public String getExcelEmailAttachFieldName() {
		return excelEmailAttachFieldName;
	}

	public void setExcelEmailAttachFieldName(String excelEmailAttachFieldName) {
		this.excelEmailAttachFieldName = excelEmailAttachFieldName;
	}

	public String getSignatureFilePath() {
		return signatureFilePath;
	}

	public void setSignatureFilePath(String signatureFilePath) {
		this.signatureFilePath = signatureFilePath;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String getExcelFilePath() {
		return excelFilePath;
	}

	public void setExcelFilePath(String excelFilePath) {
		this.excelFilePath = excelFilePath;
	}

	public String getMailBodyPath() {
		return mailBodyPath;
	}

	public void setMailBodyPath(String mailBodyPath) {
		this.mailBodyPath = mailBodyPath;
	}

	public String getMailSubject() {
		return mailSubject;
	}

	public void setMailSubject(String mailSubject) {
		this.mailSubject = mailSubject;
	}

	public String getMailServerUsername() {
		return mailServerUsername;
	}

	public void setMailServerUsername(String mailServerUsername) {
		this.mailServerUsername = mailServerUsername;
	}

	public String getMailServerHost() {
		return mailServerHost;
	}

	public void setMailServerHost(String mailServerHost) {
		this.mailServerHost = mailServerHost;
	}

	public String getMailServerPort() {
		return mailServerPort;
	}

	public void setMailServerPort(String mailServerPort) {
		this.mailServerPort = mailServerPort;
	}

	public String getMailServerPassword() {
		return mailServerPassword;
	}

	public void setMailServerPassword(String mailServerPassword) {
		this.mailServerPassword = mailServerPassword;
	}

	public int getStopEvery() {
		return stopEvery;
	}

	public void setStopEvery(int stopEvery) {
		this.stopEvery = stopEvery;
	}

	public int getStopFor() {
		return stopFor;
	}

	public void setStopFor(int stopFor) {
		this.stopFor = stopFor;
	}

	public Logger getLog() {
		return log;
	}

	public static void setInstance(ApplicationProperties instance) {
		ApplicationProperties.instance = instance;
	}

	public String[] getColHeaders() {
		return colHeaders;
	}

	public void setColHeaders(String[] colHeaders) {
		this.colHeaders = colHeaders;
	}

	public boolean isUseAuth() {
		return useAuth;
	}

	public void setUseAuth(boolean useAuth) {
		this.useAuth = useAuth;
	}	

}