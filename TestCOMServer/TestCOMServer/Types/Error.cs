using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000017"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestError
    {
        [return: MarshalAs(UnmanagedType.Error)]
        int ErrorField
        {
            [return: MarshalAs(UnmanagedType.Error)]
            get;
            [param: MarshalAs(UnmanagedType.Error)]
            set;
        }

        int PutError([MarshalAs(UnmanagedType.Error)] int value);

        [return: MarshalAs(UnmanagedType.Error)]
        int GetError();

        [return: MarshalAs(UnmanagedType.Error)]
        int EchoError([MarshalAs(UnmanagedType.Error)] int value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000018"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Error"),
        ComSourceInterfaces(typeof(ICOMTestError))]
    public class COMTestError : ICOMTestError
    {
        protected int rawError;

        public COMTestError()
        {
            rawError = 0;
        }

        public int ErrorField
        {
            get { return rawError; }
            set { rawError = value; }
        }

        public int PutError(int value) { rawError = value; return 1; }
        public int GetError() { return rawError; }
        public int EchoError(int value) { return value; }
    }
}
