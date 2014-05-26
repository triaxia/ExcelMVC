namespace ExcelMvc.Runtime
{
    using System;
    using System.Collections.Generic;

    [Serializable]
    public class Result
    {
        #region Properties

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

        #endregion Properties
    }
}
