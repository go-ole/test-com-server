using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000011"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestCurrency
    {
        [return: MarshalAs(UnmanagedType.Currency)]
        decimal CurrencyField
        {
            [return: MarshalAs(UnmanagedType.Currency)]
            get;
            [param: MarshalAs(UnmanagedType.Currency)]
            set;
        }

        int PutCurrency([MarshalAs(UnmanagedType.Currency)] decimal value);

        [return: MarshalAs(UnmanagedType.Currency)]
        decimal GetCurrency();

        [return: MarshalAs(UnmanagedType.Currency)]
        decimal EchoCurrency([MarshalAs(UnmanagedType.Currency)] decimal value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000012"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Currency"),
        ComSourceInterfaces(typeof(ICOMTestCurrency))]
    public class COMTestCurrency : ICOMTestCurrency
    {
        protected decimal rawCurrency;

        public COMTestCurrency()
        {
            rawCurrency = 0m;
        }

        public decimal CurrencyField
        {
            get { return rawCurrency; }
            set { rawCurrency = value; }
        }

        public int PutCurrency(decimal value) { rawCurrency = value; return 1; }
        public decimal GetCurrency() { return rawCurrency; }
        public decimal EchoCurrency(decimal value) { return value; }
    }
}
