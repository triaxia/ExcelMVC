﻿using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
	/// <summary>
	/// Decorates arguments of exported functions.
	/// </summary>
	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ExcelArgumentAttribute : Attribute
	{
		public string Name;
		public string Description;
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct ExcelArgument
	{
		[MarshalAs(UnmanagedType.LPWStr)]
		public string Name;
		[MarshalAs(UnmanagedType.LPWStr)]
		public string Description;

		public ExcelArgument(ExcelArgumentAttribute rhs)
        {
			Name = rhs.Name;
			Description = rhs.Description;
		}
	}
}