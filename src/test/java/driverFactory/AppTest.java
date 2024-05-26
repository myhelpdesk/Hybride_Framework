package driverFactory;

import org.testng.annotations.Test;

public class AppTest {
	@Test
	public void kickStart() throws Throwable
	{
		DrverScript ds = new DrverScript();
		ds.startTest();
	}

}
