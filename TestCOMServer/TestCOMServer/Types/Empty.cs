using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001F"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestEmpty
    {
        object GetEmpty();
        object GetNull();
        bool IsEmpty(object value);
        bool IsNull(object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000020"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Empty"),
        ComSourceInterfaces(typeof(ICOMTestEmpty))]
    public class COMTestEmpty : ICOMTestEmpty
    {
        public COMTestEmpty() { }

        public object GetEmpty() { return null; }
        public object GetNull() { return DBNull.Value; }
        public bool IsEmpty(object value) { return value == null; }
        public bool IsNull(object value) { return value is DBNull; }
    }
}
