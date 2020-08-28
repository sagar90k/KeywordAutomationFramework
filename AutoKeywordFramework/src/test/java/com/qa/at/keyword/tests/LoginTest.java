package com.qa.at.keyword.tests;

import org.junit.Test;

import com.qa.at.keyword.engine.KeywordEngine;

public class LoginTest {
	
public KeywordEngine keywordengine;
	
	@Test
	public void loginTest()
	{
	keywordengine = new KeywordEngine();
	keywordengine.startExecution("login","scenarios");
	}
	
}
