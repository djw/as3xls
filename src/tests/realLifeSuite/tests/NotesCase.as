package tests.realLifeSuite.tests {
	import com.as3xls.xls.ExcelFile;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class NotesCase extends TestCase {
		[Embed(source='/samples/notes.xls', mimeType='application/octet-stream')]
		private var b:Class;
		
		private var xls:ExcelFile;
		
		override protected function setUp():void {
			xls = new ExcelFile();
			xls.loadFromByteArray(new b());
		}
		
		public function testSheet1():void {
			assertEquals("Note 1 on Sheet1", xls.sheets[0].getCell(0, 0).note);
			assertEquals("Note 2 on Sheet1", xls.sheets[0].getCell(1, 0).note);
			assertEquals("Note 3 on Sheet1", xls.sheets[0].getCell(2, 0).note);
		}
		
		public function testSheet2():void {
			assertEquals("Note 1 on Sheet2", xls.sheets[1].getCell(0, 0).note);
			assertEquals("Note 2 on Sheet2", xls.sheets[1].getCell(1, 0).note);
			assertEquals("Note 3 on Sheet2", xls.sheets[1].getCell(2, 0).note);
		}
		
		public function testSheet3():void {
			assertEquals("Note 1 on Sheet3", xls.sheets[2].getCell(0, 0).note);
			assertEquals("Note 2 on Sheet3", xls.sheets[2].getCell(1, 0).note);
			assertEquals("Note 3 on Sheet3", xls.sheets[2].getCell(2, 0).note);
			
			assertEquals("Other Column", xls.sheets[2].getCell(0, 1).note);
			assertEquals("More other column", xls.sheets[2].getCell(0, 2).note);
		}
	}
}