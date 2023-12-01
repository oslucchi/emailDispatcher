package it.l_soft.EmailDispatcher;

public class EmailDispatcher 
{
	public static void main(String[] args) {
		System.out.println("==> started");
		new SendEmailWithAttachment(args);
	}

}
