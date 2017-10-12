package com.qlk.test;

public class Test {

	public static void main(String[] args) {
		Request request=new Request();
		System.out.println(request.getRequestURL(1));
		System.out.println(request.sendGet(1));

	}

}
