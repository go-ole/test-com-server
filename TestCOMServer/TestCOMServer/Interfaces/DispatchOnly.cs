using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000005"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICOMTestDispatchOnly
    {
        int DispatchGetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000006"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.DispatchOnly"),
        ComSourceInterfaces(typeof(ICOMTestDispatchOnly))]
    public class COMTestDispatchOnly : ICOMTestDispatchOnly
    {
        public COMTestDispatchOnly() { }

        public int DispatchGetValue() { return 126; }
    }
}
