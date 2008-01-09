package tests.utilsSuite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.utilsSuite.tests.FormatterCase;
	
	public class UtilsSuite extends TestSuite
	{
		public function UtilsSuite() {
			addTestCase(new FormatterCase());
		}

	}
}