package tests.realLifeSuite.tests
{
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	import net.digitalprimates.flex2.uint.tests.TestCase;
	public class CarleaseCase extends TestCase {
		[Embed(source='/samples/carlease.xls', mimeType='application/octet-stream')]
		private var b:Class;
		
		private var xls:ExcelFile;
		private var s:Sheet;
		
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new b());
			s = xls.sheets.getItemAt(0) as Sheet;
		}
		
		public function testFormulaValues():void {
			assertEquals(0.09380689767098427, s.getCell(15, 3).value);
			assertEquals(1276.9787527384556, s.getCell(24, 5).value);
			assertEquals(640, s.getCell(34, 4).value);
		}
		
		public function testFormulas():void {
			assertEquals("=+(1+D15/12)^12-1", s.getCell(15, 3).formula.formula);
			assertEquals("=0.06*(D11-D14)+D17*D14/(1+D15/12)^D13", s.getCell(24, 5).formula.formula);
			assertEquals("=+SUM(C35:D35)", s.getCell(34, 4).formula.formula);
		}
		
		public function testChangeValues():void {
			s.setCell(10, 3, 44000);
			s.setCell(14, 3, 0.02);
			assertEquals(26304.854397944328, s.getCell(26, 3).value);
			assertEquals(0.02018435568150201, s.getCell(15, 3).value);
			assertEquals(2570.173884876721, s.getCell(24, 5).value);
			assertEquals(1720, s.getCell(34, 4).value);
		}
	}
}