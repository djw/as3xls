package tests.cdfSuite.tests {
	import com.as3xls.cdf.CDFReader;
	
	import flash.utils.ByteArray;
	
	import net.digitalprimates.flex2.uint.tests.TestCase;
	
	public class Workbook2Case extends TestCase {
		[Embed(source='/samples/Workbook2.xls', mimeType='application/octet-stream')]
		private var work:Class;
		
		[Embed(source='/samples/biff2.xls', mimeType='application/octet-stream')]
		private var biff2:Class;
		
		private var w:ByteArray;
		
		override protected function setUp():void {
			w = new work();
		}
		
		public function testFindStreams():void {
			var cdf:CDFReader = new CDFReader(w); 
			assertEquals("Workbook", cdf.dir[0].name);
		}
		
		public function testFindWorkbook():void {
			var cdf:CDFReader = new CDFReader(w);
			var workbook:ByteArray = cdf.loadDirectoryEntry(0);
			assertEquals(0x809, workbook.readUnsignedShort());
		}
		
		public function testIsCDFFile():void {
			assertTrue(CDFReader.isCDFFile(w));
			assertFalse(CDFReader.isCDFFile(new biff2()));
		}
	}
}