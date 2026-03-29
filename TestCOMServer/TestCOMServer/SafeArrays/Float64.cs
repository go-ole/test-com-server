using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayFloat64
    {
        double[] EchoFloat64Array1D(double[] value);
        double[,] EchoFloat64Array2D(double[,] value);
        double[,,] EchoFloat64Array3D(double[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Float64"),
        ComSourceInterfaces(typeof(ICOMSafeArrayFloat64))]
    public class COMSafeArrayFloat64 : ICOMSafeArrayFloat64
    {
        public COMSafeArrayFloat64() { }

        public double[] EchoFloat64Array1D(double[] value) { return value; }
        public double[,] EchoFloat64Array2D(double[,] value) { return value; }
        public double[,,] EchoFloat64Array3D(double[,,] value) { return value; }
    }
}
