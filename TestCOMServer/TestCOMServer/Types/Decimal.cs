using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000015"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDecimal
    {
        object DecimalField { get; set; }
        int PutDecimal(object value);
        object GetDecimal();
        object EchoDecimal(object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000016"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Decimal"),
        ComSourceInterfaces(typeof(ICOMTestDecimal))]
    public class COMTestDecimal : ICOMTestDecimal
    {
        protected decimal rawDecimal;

        public COMTestDecimal()
        {
            rawDecimal = 0m;
        }

        public object DecimalField
        {
            get { return rawDecimal; }
            set { rawDecimal = (decimal)value; }
        }

        public int PutDecimal(object value) { rawDecimal = (decimal)value; return 1; }
        public object GetDecimal() { return rawDecimal; }
        public object EchoDecimal(object value) { return value; }
    }
}
