using System;

namespace TestLib
{
    public class TestClass
    {
        public string Hello()
        {
            return XX();
        }

        private string XX()
        {
            return $"{DateTime.Now:O}";
        }
    }
}
