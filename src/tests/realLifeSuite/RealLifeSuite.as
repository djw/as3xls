package tests.realLifeSuite {
	import net.digitalprimates.flex2.uint.tests.TestSuite;
	
	import tests.realLifeSuite.tests.CarleaseCase;
	import tests.realLifeSuite.tests.FormulaCase;
	import tests.realLifeSuite.tests.NotesCase;
	import tests.realLifeSuite.tests.TodoCase;
	
	public class RealLifeSuite extends TestSuite {
		public function RealLifeSuite() {
			addTestCase(new NotesCase());
			addTestCase(new CarleaseCase());
			addTestCase(new TodoCase());
			addTestCase(new FormulaCase());
		}

	}
}