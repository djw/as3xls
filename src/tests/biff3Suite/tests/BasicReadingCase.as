package tests.biff3Suite.tests {
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class BasicReadingCase extends TestCase {
		[Embed(source='/samples/biff3.xls', mimeType='application/octet-stream')]
		private var bytes:Class 
		
		private var xf:ExcelFile;
		private var s:Sheet;
		
		override protected function setUp():void {
			xf = new ExcelFile();
			xf.loadFromByteArray(new bytes());
			s = xf.sheets[0];
		}
		
		public function testGeneral():void {
			assertEquals("&LHeader\r", s.header);
			assertEquals("Footer", s.footer);
			assertEquals(10, s.rows);
			assertEquals(5, s.cols);
		}
		
		public function testIntegers():void {
			assertEquals(1, s.getCell(0, 0));
			assertEquals(2, s.getCell(1, 0));
			assertEquals(3, s.getCell(2, 0));
			assertEquals(4, s.getCell(3, 0));
			assertEquals(5, s.getCell(4, 0));
			assertEquals(6, s.getCell(5, 0));
			assertEquals(7, s.getCell(6, 0));
			assertEquals(8, s.getCell(7, 0));
			assertEquals(9, s.getCell(8, 0));
			assertEquals(10, s.getCell(9, 0));
		}
		
		public function testStrings():void {
			assertEquals("Hi", s.getCell(0, 1));
			assertEquals("you doing", s.getCell(3, 1));
			assertEquals("this merry", s.getCell(5, 1));
			assertEquals("today?", s.getCell(9, 1));
			assertEquals("Goog", s.getCell(6, 2));
		}
		
		public function testDates():void {
			assertEquals("Thu Feb 22 00:00:00 GMT-0800 2007", s.getCell(1, 2));
			assertEquals("Sat Jan 19 00:00:00 GMT-0800 2002", s.getCell(2, 2));
		}

		public function testEvalFormula():void {
			assertEquals(14.754317602356753, s.getCell(0, 3).value);
			assertEquals(14.754317602356753, s.getCell(5, 3).value);
		}
		
		public function testEvalFormulaRef():void {
			assertEquals(21.04107572533686, s.getCell(0, 4).value);
			assertEquals(131.04107572533687, s.getCell(5, 4).value);
		}
	}
}