package tests.workbookSuite.tests {
	import com.as3xls.xls.*;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class Workbook2Case extends TestCase {
		[Embed(source='/samples/Workbook2.xls', mimeType='application/octet-stream')]
		private var work:Class;
		
		
		private var xls:ExcelFile;
		private var s:Sheet;
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new work());
			s = xls.sheets[0];
		}
	
		public function testNumbers():void {
			assertEquals(47592, s.getCell(2, 0));
		}
		
		public function testStrings():void {
			assertEquals("Hi", s.getCell(0, 0));
			assertEquals("Dee", s.getCell(1, 0));
		}
		
		public function testSheets():void {
			assertEquals(3, xls.sheets.length);
			assertEquals("Sheet1", (xls.sheets[0] as Sheet).name);
			assertEquals("Sheet2", (xls.sheets[1] as Sheet).name);
			assertEquals("Sheet3", (xls.sheets[2] as Sheet).name);
		}
	}
}