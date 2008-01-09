package tests.realLifeSuite.tests {
	import com.as3xls.xls.ExcelFile;
	import com.as3xls.xls.Sheet;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class FormulaCase extends TestCase {
		[Embed(source='/samples/formulas.xls',mimeType='application/octet-stream')]
		private var b:Class;
		
		private var xls:ExcelFile;
		private var basic:Sheet;
		private var ref:Sheet;
		
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new b());
			basic = xls.sheets[0];
			ref = xls.sheets[1];
		}
		
		
		public function testSimple():void {
			assertEquals(101.19565217391305,basic.getCell(1,0));
			assertEquals("=5+8*12+18/92",basic.getCell(1,0).formula.formula);
			
			assertEquals(-37.08247422680412,basic.getCell(2,0));
			assertEquals("=55+2385/485-99+2",basic.getCell(2,0).formula.formula);
			
			assertEquals(111.71013161922252,basic.getCell(3,0));
			assertEquals("=5986*345/99^2-99",basic.getCell(3,0).formula.formula);
			
			assertEquals(-0.00508,basic.getCell(4,0));
			assertEquals("=(-1+0.984/2)%",basic.getCell(4,0).formula.formula);
		}
		
		public function testParen():void {
			             
			assertEquals("-9120.725897920605",basic.getCell(1,1));
			assertEquals("=88*(-99+(48/23)^2-9)",basic.getCell(1,1).formula.formula);
			
			assertEquals(-79.5,basic.getCell(2,1));
			assertEquals("=98-994*35/(84-99+1)^2",basic.getCell(2,1).formula.formula);
			
			assertEquals(4755.50652931135,basic.getCell(3,1));
			assertEquals("=85^2-823*(63-11*2^9)/((((57*2^-9)-85/22-(1+99)^-2.2)-84.23682*22+1)-99^-3.234)",basic.getCell(3,1).formula.formula);
		}
		
		public function testSingleArg():void {
			assertEquals(0.586121168491911,basic.getCell(1,2));
			assertEquals("=PI()*SIN(-375)+COS(99)-TAN(-1+2)+ATAN(SIN(5))",basic.getCell(1,2).formula.formula);
			
			assertEquals(877199251318.7649,basic.getCell(2,2));
			assertEquals("=SQRT(2+EXP(55)+LOG(LN(22),5)+ABS(INT(-55.482)))",basic.getCell(2,2).formula.formula);
		}
		
		public function testVarArg():void {
			assertEquals(81.87033333333333,basic.getCell(1,3));
			assertEquals("=SUM(5,482,84,-85,2.222,COUNT(5,58,8))/COUNT(5,482,84,-85,2.222,COUNT(5,58,8))",basic.getCell(1,3).formula.formula);
		}
		
		public function testSingleRef():void {
			assertEquals(-99.123,ref.getCell(17,1));
			assertEquals("=MIN(B1,B2,B3,B4,B5,B6,B7,B8,B9,B10,B11,C1:C11)",ref.getCell(17,1).formula.formula);
			
			assertEquals(7.03948573095487E+21,ref.getCell(18,1));
			assertEquals("=MAX(B1,B2,B3,B4,B5,B6,B7,B8,B9,B10,B11,C1:C11)",ref.getCell(18,1).formula.formula);
		}
		
		public function testRangeRef():void {
			assertEquals(22,ref.getCell(13,1));
			assertEquals("=COUNT(B1:C11)",ref.getCell(13,1).formula.formula);
			
			assertEquals(7.03948573095487E+21,ref.getCell(14,1));
			assertEquals("=SUM(B1:C11)",ref.getCell(14,1).formula.formula);
			
			assertEquals(319976624134312300000,ref.getCell(15,1));
			assertEquals("=SUM(B1:B11,C1:C6,C7:C11)/COUNT(B1:C11)",ref.getCell(15,1).formula.formula);
			
			assertEquals(319976624134312300000,ref.getCell(16,1));
			assertEquals("=AVERAGE(B1:C11)",ref.getCell(16,1).formula.formula);
		}
		
		public function testLogic():void {
			assertTrue(basic.getCell(1, 4).value);
			assertEquals("=AND(TRUE(),TRUE(),TRUE(),NOT(FALSE()))", basic.getCell(1, 4).formula.formula);	
		
			assertFalse(basic.getCell(2, 4).value);
			assertEquals("=OR(FALSE(),FALSE(),FALSE(),NOT(TRUE()))", basic.getCell(2, 4).formula.formula);
			
			
			assertEquals(1, basic.getCell(3, 4));
			assertEquals("=IF(5+5>=10,1,2)", basic.getCell(3, 4).formula.formula);
			
			
			assertEquals(-1, basic.getCell(4, 4));
			assertEquals("=IF(1=1,-1)", basic.getCell(4, 4).formula.formula);
			
			
			assertFalse(basic.getCell(5, 4).value);
			assertEquals('=IF(1<=0,"YES")', basic.getCell(5, 4).formula.formula);
		}
		
		public function testString():void {
			assertEquals(5, basic.getCell(1, 5));
			assertEquals('=LEN("hello")', basic.getCell(1, 5).formula.formula);
			
			
			assertEquals("TEE", basic.getCell(2, 5));
			assertEquals('=UPPER(LEFT("TEEEEXT",3))', basic.getCell(2, 5).formula.formula);
			
			
			assertEquals("EEL", basic.getCell(3, 5));
			assertEquals('=RIGHT("GOOGEEL",LEN("HOO"))', basic.getCell(3, 5).formula.formula);	
		}
		
		
	}
}