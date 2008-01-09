package tests.workbookSuite
{
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.workbookSuite.tests.Workbook1Case;
	import tests.workbookSuite.tests.Workbook2Case;
	
	public class WorkbookSuite extends TestSuite
	{
		public function WorkbookSuite() {
			addTestCase(new Workbook1Case());
			addTestCase(new Workbook2Case());
		}

	}
}