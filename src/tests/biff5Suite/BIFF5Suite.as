package tests.biff5Suite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.biff5Suite.tests.BasicReadingCase;
	import tests.biff5Suite.tests.SharedFormulaCase;
	
	public class BIFF5Suite extends TestSuite
	{
		public function BIFF5Suite() {
			addTestCase(new BasicReadingCase());
			addTestCase(new SharedFormulaCase());
		}

	}
}