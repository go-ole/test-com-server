using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFloat64
    {
        double Float64Field { get; set; }
        int PutFloat64(double value);
        double GetFloat64();
        double EchoFloat64(double value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Float64"),
        ComSourceInterfaces(typeof(ICOMTestFloat64))]
    public class COMTestFloat64 : ICOMTestFloat64
    {
        protected double rawFloat64;

        public COMTestFloat64()
        {
            rawFloat64 = 0.0;
        }

        public double Float64Field
        {
            get { return rawFloat64; }
            set { rawFloat64 = value; }
        }

        public int PutFloat64(double value) { rawFloat64 = value; return 1; }
        public double GetFloat64() { return rawFloat64; }
        public double EchoFloat64(double value) { return value; }
    }
}
