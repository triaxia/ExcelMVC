namespace ExcelMvc.Runtime
{
    using System;
    using System.Collections.Generic;

    [Serializable]
    internal class TypeResult
    {

        public Exception Error
        {
            get;
            set;
        }

        public List<string> Types
        {
            get;
            set;
        }

    }
}
