package tests.biff3Suite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.biff3Suite.tests.BasicReadingCase;
	import tests.biff3Suite.tests.FormulaCase;
	
	public class BIFF3Suite extends TestSuite
	{
		public function BIFF3Suite() {
			addTestCase(new BasicReadingCase());
			addTestCase(new FormulaCase());
		}

	}
}