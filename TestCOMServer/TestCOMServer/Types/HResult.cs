using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000023"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestHResult
    {
        int HResultField { get; set; }
        int PutHResult(int value);
        int GetHResult();
        int EchoHResult(int value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000024"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.HResult"),
        ComSourceInterfaces(typeof(ICOMTestHResult))]
    public class COMTestHResult : ICOMTestHResult
    {
        protected int rawHResult;

        public COMTestHResult()
        {
            rawHResult = 0;
        }

        public int HResultField
        {
            get { return rawHResult; }
            set { rawHResult = value; }
        }

        public int PutHResult(int value) { rawHResult = value; return 1; }
        public int GetHResult() { return rawHResult; }
        public int EchoHResult(int value) { return value; }
    }
}
