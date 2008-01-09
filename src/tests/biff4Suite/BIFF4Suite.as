package tests.biff4Suite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.biff4Suite.tests.BasicReadingCase;
	import tests.biff4Suite.tests.FormulaCase;
	
	public class BIFF4Suite extends TestSuite
	{
		public function BIFF4Suite() {
			addTestCase(new BasicReadingCase());
			addTestCase(new FormulaCase());
		}

	}
}