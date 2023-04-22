namespace ExcelMvc.Functions
{
    public enum XlTypes
    {
        xltypeNum = 0x0001,
        xltypeStr = 0x0002,
        xltypeBool = 0x0004,
        xltypeRef = 0x0008,
        xltypeErr = 0x0010,
        xltypeFlow = 0x0020,
        xltypeMulti = 0x0040,
        xltypeMissing = 0x0080,
        xltypeNil = 0x0100,
        xltypeSRef = 0x0400,
        xltypeInt = 0x0800,
        xltypeBigData = xltypeStr | xltypeInt
    }
}

