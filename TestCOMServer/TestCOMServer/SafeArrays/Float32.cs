using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayFloat32
    {
        float[] EchoFloat32Array1D(float[] value);
        float[,] EchoFloat32Array2D(float[,] value);
        float[,,] EchoFloat32Array3D(float[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Float32"),
        ComSourceInterfaces(typeof(ICOMSafeArrayFloat32))]
    public class COMSafeArrayFloat32 : ICOMSafeArrayFloat32
    {
        public COMSafeArrayFloat32() { }

        public float[] EchoFloat32Array1D(float[] value) { return value; }
        public float[,] EchoFloat32Array2D(float[,] value) { return value; }
        public float[,,] EchoFloat32Array3D(float[,,] value) { return value; }
    }
}
