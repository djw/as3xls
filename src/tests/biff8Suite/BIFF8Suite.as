package tests.biff8Suite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.biff8Suite.tests.BasicReadingCase;
	import tests.biff8Suite.tests.SharedFormulaCase;
	
	public class BIFF8Suite extends TestSuite
	{
		public function BIFF8Suite() {
			addTestCase(new BasicReadingCase());
			addTestCase(new SharedFormulaCase());
		}

	}
}