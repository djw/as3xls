package com.as3xls.xls {
	import mx.formatters.DateFormatter;
	import mx.formatters.NumberBase;
	import mx.formatters.NumberBaseRoundType;
	import mx.formatters.NumberFormatter;
	
	public class Formatter {
		private static const decimal:Object = {
			"0": buildNumberFormatter(8),
			"0.00": buildNumberFormatter(2),
			"#,##0": buildNumberFormatter(0),
			"#,##0.00": buildNumberFormatter(2)
		};
		
		private static const date:Object = {
			"d\\-mmm": "D-MMM",
			"d\\-mmm\\-yy": "D-MMM-YY",
			"mmm\\-yy": "MMM-YY"	
		};
		
		private static const currency:Object = {
			'"$"#,##0_);\\("$"#,##0\\)':			buildNumberFormatter(2),
			'"$"#,##0_);[Red]\\("$"#,##0\\)': 		buildNumberFormatter(2),
			'"$"#,##0.00_);\\("$"#,##0.00\\)': 		buildNumberFormatter(2),
			'"$"#,##0.00_);[Red]\\("$"#,##0.00\\)':	buildNumberFormatter(2)
		};
	
		private static const percent:Object = {
			"0%":"0%",
			"0.00%": "0.00%"
		};
		
		private static const scientific:Object = {
			"0.00E+00": buildNumberFormatter(-1),
			"##0.0E+0": buildNumberFormatter(-1)
		}
		
		private static function buildNumberFormatter(precision:Number):NumberFormatter {
			var ret:NumberFormatter = new NumberFormatter();
			ret.precision = precision;
			ret.useThousandsSeparator = false;
			ret.rounding = NumberBaseRoundType.NEAREST;
			return ret;
		}
		
		public static const builtInFormats:Array = [
			"General", 
			"0", "0.00", "#,##0", "#,##0.00", 
			"0%", "0.00%",
			"0.00E+00",
			"# ?/?", "# ?/?",
			"", "", "", "", "", "d\\-mmm\\-yy", "d\\-mmm", "mmm\\-yy", "",
			"", "", "", "",
			'"$"#,##0_);\\("$"#,##0\\)', '"$"#,##0_);[Red]\\("$"#,##0\\)', '"$"#,##0.00_);\\("$"#,##0.00\\)', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
			"", "", "",
			"##0.0E+0"
		];
		
		private static const rounder:NumberBase = new NumberBase();
	
		/**
		 * Determines whether the given format string represents a date. This is needed because Excel stores dates as numbers,
		 * so the only way to determine wether a cell contains a number or a date is by checking whether the format string
		 * is a date format string
		 *  
		 * @param fmt The format string to check
		 * @return True if the format string is a date format; false otherwise.
		 * 
		 */
		public static function isDate(fmt:String):Boolean {
			var state:uint = 0;
			var s:String = "";
			var c:String;
			var ignorable:Array = ["$","-","+","/","(",")",":"," "];
			for(var i:uint = 0; i < fmt.length; i++) {
				c = fmt.charAt(i);
				switch (state) {
					case 0:
						if (c == '"') {
							state = 1;
						} else if (["\\","_","*"].indexOf(c) != -1) {
							state = 2;
						} else if (ignorable.indexOf(c) != -1) {
							// do nothing
						} else {
							s = s + c;
						}
						break;
					case 1:
						if (c == '"') {
							state = 0;
						}
						break;
					case 2:
						state = 0;
						break;
				}
			}
			var fmt_bracketed_sub:RegExp = /\[[^]]*\]/;
			s = s.replace(fmt_bracketed_sub,"");
			
			var non_date_formats:Array = [
			    '0.00E+00',
			    '##0.0E+0',
			    'General' ,
			    'GENERAL' , // OOo Calc 1.1.4 does this.
			    'general' , // pyExcelerator 0.6.3 does this.
			    '@']

			if (non_date_formats.indexOf(s) != -1) { return false; }
			
			state = 0;
			
			var date_count:uint = 0;
			var num_count:uint = 0;
			var got_sep:Boolean = false;
			var date_chars:Array = ["y","m","d","h","s","Y","M","D","H","S"];
			var num_chars:Array = ["0","#","?"];
			
			for(var j:uint = 0; j < fmt.length; j++) {
				c = fmt.charAt(j);
				if (date_chars.indexOf(c) != -1){
					date_count += 5;
				} else if (num_chars.indexOf(c) != -1) {
					num_count += 5;
				} else if (c == ";"){
					got_sep = true;
				}
			}
			
			if (date_count && !num_count){ return true; }
			if (num_count && !date_count){ return false; }

			return date_count > num_count;
		}
	
		public static function format(value:*, fmt:String, base:uint = 1 /* ExcelFile.BASE1904 */):* {
			if((fmt.length == 0 || fmt == "General") &&
				isNaN(Number(value)) == false ) {
				fmt = builtInFormats[1];
			}
			if(decimal[fmt] is NumberFormatter) {
				var ret:String = decimal[fmt].format(value);
				if(ret.indexOf(".") != -1) {
					while(ret.charAt(ret.length-1) == "0") {
						ret = ret.substr(0, ret.length-1);
					}
				}
				return ret;
			} else if(percent[fmt] is String) {
				return value;
			} else if(date[fmt] is String) {
				var d:Date = new Date();
				d.hours = 0;
				d.minutes = 0;
				d.seconds = 0;
				d.milliseconds = 0;
				d.fullYear = base == ExcelFile.BASE1899 ? 1899 : 1904;
				d.month = base == ExcelFile.BASE1899 ? 11: 0;
				d.date = base == ExcelFile.BASE1899 ? 30 : 0;
				d.date += value + 1;
				
				var df:DateFormatter = new DateFormatter();
				df.formatString = date[fmt];
				return df.format(d);
			} else if(currency[fmt] is NumberFormatter) {
				return currency[fmt].format(value);
			} else {
				return value;
			}
		}

	}
}