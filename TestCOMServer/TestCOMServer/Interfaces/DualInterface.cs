using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDualInterface
    {
        int GetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000002"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Dual"),
        ComSourceInterfaces(typeof(ICOMTestDualInterface))]
    public class COMTestDualInterface : ICOMTestDualInterface
    {
        public COMTestDualInterface() { }

        public int GetValue() { return 42; }
    }
}
