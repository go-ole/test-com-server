using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000011"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayCurrency
    {
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[] EchoCurrencyArray1D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[,] EchoCurrencyArray2D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[,] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[,,] EchoCurrencyArray3D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000012"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Currency"),
        ComSourceInterfaces(typeof(ICOMSafeArrayCurrency))]
    public class COMSafeArrayCurrency : ICOMSafeArrayCurrency
    {
        public COMSafeArrayCurrency() { }

        public decimal[] EchoCurrencyArray1D(decimal[] value) { return value; }
        public decimal[,] EchoCurrencyArray2D(decimal[,] value) { return value; }
        public decimal[,,] EchoCurrencyArray3D(decimal[,,] value) { return value; }
    }
}
