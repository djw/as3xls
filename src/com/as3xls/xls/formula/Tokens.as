package com.as3xls.xls.formula {	
	/**
	 * Used to represent the RPN token array in formulas.
	 */
	public class Tokens {
		public static const NOTUSED:uint 	= 0x00;
		public static const tExp:uint 		= 0x01;
		public static const tTbl:uint 		= 0x02;
		public static const tAdd:uint 		= 0x03;
		public static const tSub:uint 		= 0x04;
		public static const tMul:uint 		= 0x05;
		public static const tDiv:uint 		= 0x06;
		public static const tPower:uint 	= 0x07;
		public static const tConcat:uint 	= 0x08;
		public static const tLT:uint 		= 0x09;
		public static const tLE:uint 		= 0x0A;
		public static const tEQ:uint 		= 0x0B;
		public static const tGE:uint 		= 0x0C;
		public static const tGT:uint 		= 0x0D;
		public static const tNE:uint 		= 0x0E;
		public static const tIsect:uint 	= 0x0F;
		public static const tList:uint 		= 0x10;
		public static const tRange:uint 	= 0x11;
		public static const tUplus:uint 	= 0x12;
		public static const tUminus:uint 	= 0x13;
		public static const tPercent:uint 	= 0x14;
		public static const tParen:uint 	= 0x15;
		public static const tMissArg:uint 	= 0x16;
		public static const tStr:uint 		= 0x17;
		public static const tNlr:uint 		= 0x18;
		public static const tAttr:uint 		= 0x19;
		public static const tSheet:uint 	= 0x1A;
		public static const tEndSheet:uint 	= 0x1B;
		public static const tErr:uint 		= 0x1C;
		public static const tBool:uint 		= 0x1D;
		public static const tInt:uint 		= 0x1E;
		public static const tNum:uint 		= 0x1F;
		
		public static const tFuncR:uint		= 0x21;
		public static const tFuncV:uint		= 0x41;
		public static const tFuncA:uint		= 0x61;
		
		public static const tFuncVarR:uint	= 0x22;
		public static const tFuncVarV:uint	= 0x42;
		public static const tFuncVarA:uint	= 0x62;
		
		public static const tRefR:uint		= 0x24;
		public static const tRefV:uint		= 0x44;
		public static const tRefA:uint		= 0x64;
		
		public static const tAreaR:uint		= 0x25;
		public static const tAreaV:uint		= 0x45;
		public static const tAreaA:uint		= 0x65;
		
		public static const tRefNR:uint		= 0x2C;
		public static const tRefNV:uint		= 0x4C;
		public static const tRefNA:uint		= 0x6C;
		
		
		
		
		
		
		public static const tFuncCER:uint	= 0x38;
		public static const tFuncCEV:uint	= 0x58;
		public static const tFuncCEA:uint	= 0x78;
		
		
		public static const tRef3d:uint		= 0x3A;
		
		
		
		
	}
}