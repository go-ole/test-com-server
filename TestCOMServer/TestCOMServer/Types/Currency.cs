using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000011"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestCurrency
    {
        long CurrencyField { get; set; }
        int PutCurrency(long value);
        long GetCurrency();
        long EchoCurrency(long value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000012"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Currency"),
        ComSourceInterfaces(typeof(ICOMTestCurrency))]
    public class COMTestCurrency : ICOMTestCurrency
    {
        protected long rawCurrency;

        public COMTestCurrency()
        {
            rawCurrency = 0;
        }

        public long CurrencyField
        {
            get { return rawCurrency; }
            set { rawCurrency = value; }
        }

        public int PutCurrency(long value) { rawCurrency = value; return 1; }
        public long GetCurrency() { return rawCurrency; }
        public long EchoCurrency(long value) { return value; }
    }
}
