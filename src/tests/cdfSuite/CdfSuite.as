package tests.cdfSuite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.cdfSuite.tests.Workbook2Case;
	
	public class CdfSuite extends TestSuite
	{
		public function CdfSuite() {
			addTestCase(new Workbook2Case());
		}

	}
}