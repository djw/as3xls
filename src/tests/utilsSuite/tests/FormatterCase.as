package tests.utilsSuite.tests
{
	import net.digitalprimates.flex2.uint.tests.TestCase;
	import com.as3xls.xls.Formatter;
	public class FormatterCase extends TestCase
	{
		public function testGeneral():void {
			assertEquals(12.222, Formatter.format(12.222, "General"));
		}

		public function testDate1():void {
			assertEquals("22-Feb", Formatter.format(37673, "d\\-mmm"));
		}
		
		public function testDate2():void {
			assertEquals("22-Feb-07", Formatter.format(37673, "d\\-mmm\\-yy"));
		}
	}
}