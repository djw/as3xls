package tests.biff2Suite.tests
{
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class FormulaCase extends TestCase {
		[Embed(source='/samples/biff2.xls', mimeType='application/octet-stream')]
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
			assertEquals("=A6*22+SIN(5)", s.getCell(5, 4).formula.formula);
		}
	}
}