using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFloat32
    {
        float Float32Field { get; set; }
        int PutFloat32(float value);
        float GetFloat32();
        float EchoFloat32(float value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Float32"),
        ComSourceInterfaces(typeof(ICOMTestFloat32))]
    public class COMTestFloat32 : ICOMTestFloat32
    {
        protected float rawFloat32;

        public COMTestFloat32()
        {
            rawFloat32 = 0.0f;
        }

        public float Float32Field
        {
            get { return rawFloat32; }
            set { rawFloat32 = value; }
        }

        public int PutFloat32(float value) { rawFloat32 = value; return 1; }
        public float GetFloat32() { return rawFloat32; }
        public float EchoFloat32(float value) { return value; }
    }
}
