package tests.workbookSuite.tests {
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class Workbook1Case extends TestCase {
		[Embed(source='/samples/Workbook1.xls', mimeType='application/octet-stream')]
		private var work:Class;
		
		
		private var xls:ExcelFile;
		private var s:Sheet;
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new work());
			s = xls.sheets[0];
		}
		
		public function testNumbers():void {
			assertEquals(5, s.getCell(0, 0));
			assertEquals(6, s.getCell(3, 0));
		}
		
		public function testStrings():void {
			assertEquals("Hollo", s.getCell(2, 1));
			assertEquals("Hi", s.getCell(3, 1));
		}
		
		public function testBlank():void {
			assertEquals("", s.getCell(1, 1));
		}
	}
}