using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestUnknown
    {
        object UnknownField
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            get;
            [param: MarshalAs(UnmanagedType.IUnknown)]
            set;
        }

        int PutUnknown([MarshalAs(UnmanagedType.IUnknown)] object value);

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object GetUnknown();

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object EchoUnknown([MarshalAs(UnmanagedType.IUnknown)] object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Unknown"),
        ComSourceInterfaces(typeof(ICOMTestUnknown))]
    public class COMTestUnknown : ICOMTestUnknown
    {
        protected object rawUnknown;

        public COMTestUnknown()
        {
            rawUnknown = null;
        }

        public object UnknownField
        {
            get { return rawUnknown; }
            set { rawUnknown = value; }
        }

        public int PutUnknown(object value) { rawUnknown = value; return 1; }
        public object GetUnknown() { return rawUnknown; }
        public object EchoUnknown(object value) { return value; }
    }
}
