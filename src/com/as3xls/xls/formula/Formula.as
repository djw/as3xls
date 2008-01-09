package com.as3xls.xls.formula {
	import com.as3xls.biff.BIFFVersion;
	import com.as3xls.xls.Sheet;
	
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	
	/**
	 * <p>
	 * Represents Formulas in Excel formulas and handles converting them to human readable format and evaluating them.
	 * </p>
	 * 
	 * <p>
	 * Excel stores formulas as RPN tokens. This makes it relatively easy to evaluate them but a bit of a pain to display
	 * them in human readable format. You can't please everyone, I guess.
	 * </p>
	 * 
	 */
	public class Formula {
		private var _result:*;
		private var _alwaysRecalculate:Boolean;
		private var _formula:String;
		private var _tokens:ByteArray;
		private var _sheet:Sheet;
		private var biffVer:uint;
		
		private var dirty:Boolean;
		
		public var myRow:uint;
		public var myCol:uint;
		
		private var isSharedFormula:Boolean;
		private var sharedFormulaRow:uint;
		private var sharedFormulaCol:uint;
		
		
		
		/**
		 * Creates a new Formula in a given cell with an array of tokens. Note that the row, col, and so on don't
		 * affect the actual location of the formula. They are only used when relative cell locations rear their
		 * ugly but admittedly useful heads. The location is determined by this formula's cell object
		 * 
		 * @param row Formula's cell's row
		 * @param col Formula's cell's col
		 * @param tokens A ByteArray containing the tokens in Excel's native RPN format
		 * @param biffVer The BIFF version used to encode the Formula
		 * @param sheet The sheet which the Formula inhabits
		 * 
		 * @see com.as3xls.xls.Cell
		 * 
		 */
		public function Formula(row:uint, col:uint, tokens:ByteArray, biffVer:uint, sheet:Sheet) {
			myRow = row;
			myCol = col;
			_result = result;
			_tokens = tokens;
			_tokens.endian = flash.utils.Endian.LITTLE_ENDIAN;
			this.biffVer = biffVer;
			_sheet = sheet;
			_formula = "";
			
			isSharedFormula = false;			
			dirty = true;
		}
		
		/**
		 * Sums the Numbers contained in the given array 
		 * @param args An array of Numbers to add up
		 * @return The sum of the given numbers
		 * 
		 */
		private function builtInSum(args:Array):Number {
 			var ret:Number = 0;
			for(var n:uint = 0; n < args.length; n++) {
				var num:Number = Number(args[n]);
				ret += num;
			}
			return ret;
		}
		
		/**
		 * Recursively converts the given object into a one-dimensional array. If an array of multiple dimensions is passed in
		 * the result will still be a one-dimensional array containing each element.
		 *  
		 * @param arg The object to convert
		 * @return A one-dimensional array containing the element passed in or elements of the array passed in.
		 * 
		 */
		private function convertToArray(arg:*):Array {
			var ret:Array = new Array();
			if(arg is Array) {
				for(var i:uint = 0; i < arg.length; i++) {
					ret = ret.concat(convertToArray(arg[i]));
				}
			} else {
				ret = [arg];
			}
			return ret;
		}
		
		
		/**
		 * Executes one of Excel's built in functions (obviously without relying on Excel).
		 * 
		 * @param idx The idx of the function to use
		 * @param rest The arguments of the function to be called
		 * @return The result of the operation
		 * 
		 */
		private function builtInFunction(idx:uint, ... rest):* {
			var ret:*;
			var a:Array;
			var n:Number;
			switch(idx) {
				case 0:		// COUNT
					return convertToArray(rest).length;
				case 1:		// IF
					return Boolean(rest[0]) ? rest[1] : rest[2];
				case 4:		// SUM
					ret = 0;
					a = convertToArray(rest);
					for(var i:uint = 0; i < a.length; i++) {
						ret += a[i];
					}
					return ret;
				case 5:		// AVERAGE
					a = convertToArray(rest);
					return builtInSum(a) / a.length;
				case 6:		// MIN
					return Math.min.apply(null, convertToArray(rest));
				case 7:		// MAX
					return Math.max.apply(null, convertToArray(rest));
				case 15:	// SIN
					return Math.sin(rest[0]);
				case 16:	// COS
					return Math.cos(rest[0]);
				case 17:	// TAN
					return Math.tan(rest[0]);
				case 18:	// ARCTAN
					return Math.atan(rest[0]);
				case 19:	// PI
					return Math.PI;
				case 20:	// SQRT
					return Math.sqrt(rest[0]);
				case 21:	// EXP
					return Math.exp(rest[0]);
				case 22:	// LN
					return Math.log(rest[0]) / Math.LOG10E;
				case 23:	// LOG10	
					return Math.log(rest[0]);
				case 24:	// ABS
					return Math.abs(rest[0]);
				case 25:	// INT
					return Math.round(rest[0]);
				case 30:	// REPT
					var text:String = String(rest[0]);
					var count:Number = Number(rest[1]);
					ret = "";
					for(n = 0; n < count; n++){
						ret += text;
					}
					return ret;
				case 31:	// MID
					return String(rest[0]).substr(Number(rest[1]), Number(rest[2]));
				case 32:	// LEN
					return String(rest[0]).length;
				case 34:	// TRUE
					return true;
				case 35:	// FALSE
					return false;
				case 36:	// AND
					for(n = 0; n < rest.length; n++) {
						if(Boolean(rest[n]) == false) {
							return false;
						}
					}
					return true;
				case 37:	// OR
					for(n = 0; n < rest.length; n++) {
						if(Boolean(rest[n]) == true) {
							return true;
						}
					}
					return false;
				case 38:	// NOT
					return !Boolean(rest[0]);
				case 39:	// MOD
					return Number(rest[0]) % Number(rest[1]);
				case 56:	// PV
					var rate:Number = Number(rest[0]);
					var nper:Number = Number(rest[1]);
					var pmt:Number = Number(rest[2]);
					return -pmt*((Math.pow(1+rate, nper)-1)/rate) / (Math.pow(1+rate, nper));
				case 63:	// RAND
					return Math.random();
				case 97:	// ATAN2
					return Math.atan2(Number(rest[0]), Number(rest[1]));
				case 98:	// ASIN
					return Math.asin(Number(rest[0]));
				case 99:	// ACOS
					return Math.acos(Number(rest[0]));
				case 109:	// LOG
					return Math.log(rest[0]);
				case 111:	// CHAR
					return String.fromCharCode(rest[0]);
				case 112:	// LOWER
					return String(rest[0]).toLowerCase();
				case 113:	// UPPER
					return String(rest[0]).toUpperCase();
				case 115:	// LEFT
					return String(rest[0]).substr(0, Number(rest[1]));
				case 116:	// RIGHT
					return String(rest[0]).substr(String(rest[0]).length-Number(rest[1]), Number(rest[1]));
				default:
					throw new Error("Unsupported function: " + idx);
			}
		}
		
		
		/**
		 * Evaluates the represented function and stores the result 
		 * 
		 */
		public function updateResult():void {
			if(_tokens == null) {
				return;
			}
			_tokens.position = 0;
			
			var tok:uint;
			var v1:*;
			var v2:*;
			var idx:uint;
			var n:uint;
			var numArgs:uint;
			var args:Array;
			var row:uint;
			var col:uint;
			var rowEnd:uint;
			var colEnd:uint;
			
			var stack:Array = new Array();
			var unknown:Array = new Array();
			
			while(_tokens.bytesAvailable > 0) {
				tok = _tokens.readUnsignedByte();
				
				switch(tok) {
					case Tokens.tExp:
						row = _tokens.readUnsignedShort();
						col = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
						sharedFormulaRow = row;
						sharedFormulaCol = col;
						
						if(_sheet.getCell(row, col).sharedTokens == null) {
							return;
						}
						_tokens = _sheet.getCell(row, col).sharedTokens;
						updateResult();
						return;
						break;
						
					// Constant tokens
					case Tokens.tBool:
						stack.push(_tokens.readUnsignedByte() == 1);
						break;
					case Tokens.tNum:
						stack.push(_tokens.readDouble());
						break;
					case Tokens.tInt:
						stack.push(_tokens.readUnsignedShort());
						break;
					case Tokens.tStr:
						var len:uint = _tokens.readUnsignedByte();
						if(biffVer == BIFFVersion.BIFF8) {
							_tokens.position++;	// Skip option byte
						}
						
						stack.push(_tokens.readUTFBytes(len));
						break;
						
					// Unary oporators
					case Tokens.tUplus:
						stack.push(stack.pop());
						break;
					case Tokens.tUminus:
						stack.push(stack.pop() * -1);
						break;
					case Tokens.tPercent:
						stack.push(stack.pop() / 100);
						break;
						
					// Binary operators
					case Tokens.tAdd:
						stack.push(stack.pop() + stack.pop());
						break;
					case Tokens.tSub:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 - v1);
						break;
					case Tokens.tMul:
						stack.push(stack.pop() * stack.pop());
						break;
					case Tokens.tDiv:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 / v1);
						break;
					case Tokens.tPower:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(Math.pow(v2, v1));
						break;
					case Tokens.tLT:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 < v1);
						break;
					case Tokens.tLE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 <= v1);
						break;
					case Tokens.tEQ:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 == v1);
						break;
					case Tokens.tGE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 >= v1);
						break;
					case Tokens.tGT:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 > v1);
						break;
					case Tokens.tNE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 != v1);
						break;
					case Tokens.tConcat:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + v1);
						break;
					case Tokens.tParen:
						break;
					
					// Functions
					case Tokens.tFuncR:
					case Tokens.tFuncV:
					case Tokens.tFuncA:
						idx = biffVer <= BIFFVersion.BIFF3 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
						numArgs = Functions.numArgs[idx];
						args = new Array();
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						args.push(idx);
						stack.push(builtInFunction.apply(this, args.reverse()));
						break;
					case Tokens.tFuncVarR:
					case Tokens.tFuncVarV:
					case Tokens.tFuncVarA:
						if(biffVer <= BIFFVersion.BIFF3) {
							numArgs = _tokens.readUnsignedByte();
							idx = _tokens.readUnsignedByte();
						} else {
							numArgs = _tokens.readUnsignedByte() & 0x7F;
							idx = _tokens.readUnsignedShort() & 0x7FFF;
						}
						args = new Array();
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						args.push(idx);
						stack.push(builtInFunction.apply(this, args.reverse()));
						break;
					case Tokens.tFuncCER:
					case Tokens.tFuncCEV:
					case Tokens.tFuncCEA:
						numArgs = _tokens.readUnsignedByte();
						idx = _tokens.readUnsignedByte();
						args = new Array();
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						args.push(idx);
						stack.push(builtInFunction.apply(this, args.reverse()));
						break;
						
					case Tokens.tRefNR:
					case Tokens.tRefNV:
					case Tokens.tRefNA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readShort();
							col = _tokens.readByte();
							var flag:uint = _tokens.readUnsignedByte();
							
							if((flag & 0x80) != 0) {
								row = myRow + row;
							}
							if((flag & 0x40) != 0) {
								col = myCol + col;
							}
						} else {
							row = _tokens.readShort() + 1;
							col = _tokens.readByte();
							if(row & 0x8000 != 0) {
								col = myCol + col;
							}
							if(row & 0x4000 != 0) {
								row = myRow + row;
							}
							row &= 0x3FFF;
						}
						
						// HACK Seems BIFF5 is off by one???
						if(biffVer == BIFFVersion.BIFF5) {
							row--;
						}
						
						stack.push(_sheet.getCell(row, col).value);
						break;
					case Tokens.tRefR:
					case Tokens.tRefV:
					case Tokens.tRefA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readUnsignedShort();
							col = _tokens.readUnsignedShort() & 0x00FF;
						} else {
							row = _tokens.readShort() & 0x3FFF;
							col = _tokens.readByte();
						}
						
						stack.push(_sheet.getCell(row, col).value);
						break;
					case Tokens.tAreaR:
					case Tokens.tAreaV:
					case Tokens.tAreaA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readUnsignedShort();
							rowEnd = _tokens.readUnsignedShort(); 
							col = _tokens.readUnsignedShort() & 0x00FF;
							colEnd = _tokens.readUnsignedShort() & 0x00FF;
						} else {
							row = _tokens.readUnsignedShort() & 0x3FFF;
							rowEnd = _tokens.readUnsignedShort() & 0x3FFF; 
							col = _tokens.readUnsignedByte();
							colEnd = _tokens.readUnsignedByte();
						}
						
						var cells:Array = new Array();
						for(var r:uint = row; r <= rowEnd; r++) {
							cells[r-row] = new Array();
							for(var c:uint = col; c <= colEnd; c++) {
								cells[r-row][c-col] = _sheet.getCell(r, c).value;
							}
						}
						stack.push(cells);
						break;
					case Tokens.tAttr:
						var attrType:uint = _tokens.readUnsignedByte();
						var dist:uint;
						switch(attrType) {
							case 0x02:	// tAttrIF: IF Token control
								dist = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
								if(stack.pop() == false) {
									_tokens.position += dist;
								}
								break;
							case 0x08:	// tAttrSkip: Jump-like command
								dist = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte()+1 : _tokens.readUnsignedShort()+1;
								_tokens.position += dist;
								break;
							case 0x10:	// tAttrSum: SUM with 1 parameter
								_tokens.position += biffVer == BIFFVersion.BIFF2 ? 1 : 2;
								stack.push(builtInSum(convertToArray(stack.pop())));
								break;
							case 0x40: // tAttrSpace: Need to skip an extra space
								_tokens.position += biffVer == BIFFVersion.BIFF2 ? 1 : 2;
								break;
						}
						break;
					
					default:
						unknown.push(tok);
						break;
				}
				
			}
			if(unknown.length > 0) {
				throw new Error("Unknown formula tokens: " + unknown.join(", "));
			}
			_result = stack.pop();
		}
		
		
		
		/**
		 * Updates the human readable representation of the formula 
		 */
		public function updateHumanReadable():void {
			_tokens.position = 0;
			
			var tok:uint;
			var v1:*;
			var v2:*;
			var idx:uint;
			var n:uint;
			var numArgs:uint;
			var args:Array;
			var row:int;
			var col:int;
			var rowEnd:int;
			var colEnd:int;
			
			var stack:Array = new Array();
			var unknown:Array = new Array();
			
			while(_tokens.bytesAvailable > 0) {
				tok = _tokens.readUnsignedByte();
				
				switch(tok) {
					case Tokens.tExp:
						row = _tokens.readUnsignedShort();
						col = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
						sharedFormulaRow = row;
						sharedFormulaCol = col;
						
						if(_sheet.getCell(row, col).sharedTokens == null) {
							return;
						}
						_tokens = _sheet.getCell(row, col).sharedTokens;
						updateHumanReadable();
						return;
						break;
						
						
					// Constant Operators
					case Tokens.tBool:
						v1 = _tokens.readUnsignedByte() == 1;
						stack.push(v1.toString().toUpperCase());
						break;
					case Tokens.tNum:
						stack.push(_tokens.readDouble());
						break;
					case Tokens.tInt:
						stack.push(_tokens.readUnsignedShort());
						break;
					case Tokens.tStr:
						var len:uint = _tokens.readUnsignedByte();
						if(biffVer == BIFFVersion.BIFF8) {
							_tokens.position++;	// Skip option byte
						}
						
						stack.push('"' + _tokens.readUTFBytes(len) + '"');
						break;
						
					// Unary Operators
					case Tokens.tUplus:
						stack.push("+" + stack.pop());
						break;
					case Tokens.tUminus:
						stack.push("-" + stack.pop());
						break;
					case Tokens.tPercent:
						stack.push(stack.pop() + "%");
						break;
					
						
					// Binary Operators	
					case Tokens.tAdd:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "+" + v1);
						break;
					case Tokens.tSub:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "-" + v1);
						break;
					case Tokens.tMul:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "*" + v1);
						break;
					case Tokens.tDiv:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "/" + v1);
						break;
					case Tokens.tPower:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "^" + v1);
						break;
					case Tokens.tConcat:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push('"' + v2 + '"' + "&" + '"' + v1 + '"');
						break;
						
						
					case Tokens.tParen:
						stack.push("(" + stack.pop() + ")");
						break;
						
						
					// Logic operators
					case Tokens.tLT:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "<" + v1);
						break;
					case Tokens.tLE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "<=" + v1);
						break;
					case Tokens.tEQ:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "=" + v1);
						break;
					case Tokens.tGE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + ">=" + v1);
						break;
					case Tokens.tGT:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + ">" + v1);
						break;
					case Tokens.tNE:
						v1 = stack.pop();
						v2 = stack.pop();
						stack.push(v2 + "!=" + v1);
						break;
					
					// Function operators
					case Tokens.tFuncR:
					case Tokens.tFuncV:
					case Tokens.tFuncA:
						idx = biffVer <= BIFFVersion.BIFF3 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
						args = new Array();
						numArgs = Functions.numArgs[idx];
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						stack.push(Functions.names[idx] + "(" + args.join(",") + ")");
						break;
					case Tokens.tFuncVarR:
					case Tokens.tFuncVarV:
					case Tokens.tFuncVarA:
						if(biffVer <= BIFFVersion.BIFF3) {
							numArgs = _tokens.readUnsignedByte();
							idx = _tokens.readUnsignedByte();
						} else {
							numArgs = _tokens.readUnsignedByte() & 0x7F;
							idx = _tokens.readUnsignedShort() & 0x7FFF;
						}
						args = new Array();
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						stack.push(Functions.names[idx] + "(" + args.reverse().join(",") + ")");
						break;
					case Tokens.tFuncCER:
					case Tokens.tFuncCEV:
					case Tokens.tFuncCEA:
						numArgs = _tokens.readUnsignedByte();
						idx = _tokens.readUnsignedByte();
						args = new Array();
						for(n = 0; n < numArgs; n++) {
							args.push(stack.pop());
						}
						stack.push(Functions.names[idx] + "(" + args.reverse().join(",") + ")");
						break;
					case Tokens.tRefNR:
					case Tokens.tRefNV:
					case Tokens.tRefNA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readShort();
							col = _tokens.readByte();
							var flag:uint = _tokens.readUnsignedByte();
							
							if((flag & 0x80) != 0) {
								row = myRow + row;
							}
							if((flag & 0x40) != 0) {
								col = myCol + col;
							}
						} else {
							row = _tokens.readShort() + 1;
							col = _tokens.readByte();
							if(row & 0x8000 != 0) {
								col = myCol + col;
							}
							if(row & 0x4000 != 0) {
								row = myRow + row;
							}
							row &= 0x3FFF;
						}
						
						// HACK Seems BIFF5 is off by one???
						if(biffVer == BIFFVersion.BIFF5) {
							row--;
						}
						
						stack.push(String.fromCharCode(col + 0x41) + (row+1));
						break;
					case Tokens.tRefR:
					case Tokens.tRefV:
					case Tokens.tRefA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readUnsignedShort();
							col = _tokens.readUnsignedShort() & 0x00FF;
						} else {
							row = _tokens.readShort() & 0x3FFF;
							col = _tokens.readByte();
						}
						
						stack.push(String.fromCharCode(col + 0x41) + (row+1));
						break;
					case Tokens.tAreaR:
					case Tokens.tAreaV:
					case Tokens.tAreaA:
						if(biffVer == BIFFVersion.BIFF8) {
							row = _tokens.readUnsignedShort();
							rowEnd = _tokens.readUnsignedShort(); 
							col = _tokens.readUnsignedShort() & 0x00FF;
							colEnd = _tokens.readUnsignedShort() & 0x00FF;
						} else {
							row = _tokens.readUnsignedShort() & 0x3FFF;
							rowEnd = _tokens.readUnsignedShort() & 0x3FFF; 
							col = _tokens.readUnsignedByte();
							colEnd = _tokens.readUnsignedByte();
						}
					
						stack.push(String.fromCharCode(col + 0x41) + (row+1) + ":" + String.fromCharCode(colEnd + 0x41) + (rowEnd+1));
						break;
					
					case Tokens.tAttr:
						var attrType:uint = _tokens.readUnsignedByte();
						var dist:uint;
						switch(attrType) {
							case 0x02:	// tAttrIF: IF Token control
								dist = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
								break;
							case 0x08:	// tAttrSkip: Jump-like command
								dist = biffVer == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte()+1 : _tokens.readUnsignedShort()+1;
								break;
							case 0x10:	// tAttrSum: SUM with 1 parameter
								_tokens.position += biffVer == BIFFVersion.BIFF2 ? 1 : 2;
								stack.push("SUM(" + stack.pop() + ")");
								break;
							case 0x40: // tAttrSpace: Need to skip an extra space
								_tokens.position += biffVer == BIFFVersion.BIFF2 ? 1 : 2;
								break;
						}
						break;
					default:
						unknown.push(tok);
						break;
				}
				
			}
			if(unknown.length > 0) {
				throw new Error("Unknown formula tokens: " + unknown.join(", "));
			}
			_formula = "=" + stack.pop();
			dirty = false;
		}
		
		
		/**
		 * 
		 * @return The human-readable version of the formula 
		 * 
		 */
		public function get formula():String { if(dirty) { updateHumanReadable(); } return _formula; }

		/**
		 * 
		 * @return The result of evaluating the formula 
		 * 
		 */
		public function get result():* { updateResult(); return _result; }
		public function set result(value:*):void { _result = value; }
	}
}