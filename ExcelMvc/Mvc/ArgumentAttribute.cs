using System;
using System.Runtime.InteropServices;

namespace Mvc
{
	/// <summary>
	/// Decorates arguments of exported functions.
	/// </summary>
	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ArgumentAttribute : Attribute
	{
		public string Name;
		public string Description;
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct Argument
	{
		[MarshalAs(UnmanagedType.LPWStr)]
		public string Name;
		[MarshalAs(UnmanagedType.LPWStr)]
		public string Description;

		public Argument(ArgumentAttribute rhs)
        {
			Name = rhs.Name;
			Description = rhs.Description;
		}
	}
}
