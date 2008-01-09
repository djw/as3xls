package tests.realLifeSuite.tests
{
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	public class TodoCase extends TestCase {
		[Embed(source='/samples/to do 2007.xls', mimeType='application/octet-stream')]
		private var b:Class;
		
		private var xls:ExcelFile;
		private var s:Sheet;
		
		
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new b());
			s = xls.sheets.getItemAt(0) as Sheet;
		}
		
		public function testSheets():void {
			assertEquals(12, xls.sheets.length);
			assertEquals("Dec 07", xls.sheets.getItemAt(0).name);
		}
		
		public function testText():void {
			assertEquals("Paperwork", s.getCell(2, 2));
			assertEquals("Caulking", s.getCell(7, 2));
			assertEquals("Ext. caulking", s.getCell(11, 2));
			assertEquals("Outside walls", s.getCell(15, 2));
			assertEquals("Up. Hall", s.getCell(30, 2));
		}
		
		public function testMoreText():void {
			assertEquals("", s.getCell(5, 4));
			assertEquals("12.16", s.getCell(9, 4));
			assertEquals("blind", s.getCell(26, 11));
		}
		
		public function testOtherSheet():void {
			s = xls.sheets.getItemAt(4) as Sheet;
			assertEquals(8.7, s.getCell(11, 3).value);
			assertEquals("computer backup", s.getCell(29, 2));
		}
		
		public function testValues():void {
			assertEquals(2, s.getCell(2, 8).value);
		}
		
		public function testFormulasEval():void {
			assertEquals(17, s.getCell(32, 8));
			assertEquals(27, s.getCell(32, 9));
		}
		
		public function testFormulas():void {
			assertEquals("=SUM(I3:I32)", s.getCell(32, 8).formula.formula);
			assertEquals("=SUM(J3:J32)", s.getCell(32, 9).formula.formula);
		}
	}
}