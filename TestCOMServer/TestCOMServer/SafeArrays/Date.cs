using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000013"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayDate
    {
        DateTime[] EchoDateArray1D(DateTime[] value);
        DateTime[,] EchoDateArray2D(DateTime[,] value);
        DateTime[,,] EchoDateArray3D(DateTime[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000014"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Date"),
        ComSourceInterfaces(typeof(ICOMSafeArrayDate))]
    public class COMSafeArrayDate : ICOMSafeArrayDate
    {
        public COMSafeArrayDate() { }

        public DateTime[] EchoDateArray1D(DateTime[] value) { return value; }
        public DateTime[,] EchoDateArray2D(DateTime[,] value) { return value; }
        public DateTime[,,] EchoDateArray3D(DateTime[,,] value) { return value; }
    }
}
