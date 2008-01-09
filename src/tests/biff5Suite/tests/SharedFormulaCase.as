package tests.biff5Suite.tests
{
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class SharedFormulaCase extends TestCase {
		[Embed(source='/samples/biff5.xls', mimeType='application/octet-stream')]
		private var bytes:Class 
		
		private var xf:ExcelFile;
		private var s:Sheet;
		
		override protected function setUp():void {
			xf = new ExcelFile();
			xf.loadFromByteArray(new bytes());
			s = xf.sheets[0];
		}

		public function testFormula():void {
			assertEquals("=5+15*SIN(15)", s.getCell(0, 3).formula.formula);
			assertEquals("=5+15*SIN(15)", s.getCell(5, 3).formula.formula);
		}
		
		public function testFormulaRef():void {
			assertEquals("=A1*22+SIN(5)", s.getCell(0, 4).formula.formula);
			assertEquals("=A2*22+SIN(5)", s.getCell(1, 4).formula.formula);
			assertEquals("=A3*22+SIN(5)", s.getCell(2, 4).formula.formula);
			assertEquals("=A4*22+SIN(5)", s.getCell(3, 4).formula.formula);
			assertEquals("=A5*22+SIN(5)", s.getCell(4, 4).formula.formula);
			assertEquals("=A6*22+SIN(5)", s.getCell(5, 4).formula.formula);
			assertEquals("=A7*22+SIN(5)", s.getCell(6, 4).formula.formula);
			assertEquals("=A8*22+SIN(5)", s.getCell(7, 4).formula.formula);
			assertEquals("=A9*22+SIN(5)", s.getCell(8, 4).formula.formula);
			assertEquals("=A10*22+SIN(5)", s.getCell(9, 4).formula.formula);
		}
	}
}