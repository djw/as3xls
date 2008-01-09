package tests.biff2Suite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.biff2Suite.tests.BasicReadingCase;
	import tests.biff2Suite.tests.FormulaCase;
	
	public class BIFF2Suite extends TestSuite
	{
		public function BIFF2Suite() {
			addTestCase(new BasicReadingCase());
			addTestCase(new FormulaCase());
		}

	}
}