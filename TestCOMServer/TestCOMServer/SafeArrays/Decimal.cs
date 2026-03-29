using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000015"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayDecimal
    {
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[] EchoDecimalArray1D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[,] EchoDecimalArray2D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[,] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[,,] EchoDecimalArray3D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000016"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Decimal"),
        ComSourceInterfaces(typeof(ICOMSafeArrayDecimal))]
    public class COMSafeArrayDecimal : ICOMSafeArrayDecimal
    {
        public COMSafeArrayDecimal() { }

        public decimal[] EchoDecimalArray1D(decimal[] value) { return value; }
        public decimal[,] EchoDecimalArray2D(decimal[,] value) { return value; }
        public decimal[,,] EchoDecimalArray3D(decimal[,,] value) { return value; }
    }
}
