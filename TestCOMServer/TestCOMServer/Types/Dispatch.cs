using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001D"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDispatch
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object DispatchField
        {
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
            [param: MarshalAs(UnmanagedType.IDispatch)]
            set;
        }

        int PutDispatch([MarshalAs(UnmanagedType.IDispatch)] object value);

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object GetDispatch();

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object EchoDispatch([MarshalAs(UnmanagedType.IDispatch)] object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Dispatch"),
        ComSourceInterfaces(typeof(ICOMTestDispatch))]
    public class COMTestDispatch : ICOMTestDispatch
    {
        protected object rawDispatch;

        public COMTestDispatch()
        {
            rawDispatch = null;
        }

        public object DispatchField
        {
            get { return rawDispatch; }
            set { rawDispatch = value; }
        }

        public int PutDispatch(object value) { rawDispatch = value; return 1; }
        public object GetDispatch() { return rawDispatch; }
        public object EchoDispatch(object value) { return value; }
    }
}
