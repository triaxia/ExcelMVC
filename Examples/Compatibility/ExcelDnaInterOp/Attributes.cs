using Function.Interfaces;
using System;

namespace ExcelDnaInterOp
{
    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class FunctionAttribute : ExcelDna.Integration.ExcelFunctionAttribute, IFunctionAttribute
    {
        public bool IsAsync { get; set; }
        public new string Category { get => base.Category; set => base.Category = value; }
        public new string Name { get => base.Name; set => base.Name = value; }
        public new string Description { get => base.Description; set => base.Description = value; }
        public new string HelpTopic { get => base.HelpTopic; set => base.HelpTopic = value; }
        public new bool IsVolatile { get => base.IsVolatile; set => base.IsVolatile = value; }
        public new bool IsMacroType { get => base.IsMacroType; set => base.IsMacroType = value; }
        public new bool IsHidden { get => base.IsHidden; set => base.IsHidden = value; }
        public new bool IsThreadSafe { get => base.IsThreadSafe; set => base.IsThreadSafe = value; }
        public new bool IsClusterSafe { get => base.IsClusterSafe; set => base.IsClusterSafe = value; }
        public FunctionAttribute() { }
        public FunctionAttribute(string description) => base.Description = description;
    }

    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public class ArgumentAttribute : ExcelDna.Integration.ExcelArgumentAttribute, IArgumentAttribute
    {
        public new string Name { get => base.Name; set => base.Name = value; }
        public new string Description { get => base.Description; set => base.Description = value; }
        public ArgumentAttribute() { }
        public ArgumentAttribute(string description) => base.Description = description;
    }
}
