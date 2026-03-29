using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000002A"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestBlob
    {
        byte[] BlobField { get; set; }
        int PutBlob(byte[] value);
        byte[] GetBlob();
        byte[] EchoBlob(byte[] value);
        byte[] CreateBlob(int size);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000002B"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Blob"),
        ComSourceInterfaces(typeof(ICOMTestBlob))]
    public class COMTestBlob : ICOMTestBlob
    {
        protected byte[] rawBlob;

        public COMTestBlob()
        {
            rawBlob = null;
        }

        public byte[] BlobField
        {
            get { return rawBlob; }
            set { rawBlob = value; }
        }

        public int PutBlob(byte[] value) { rawBlob = value; return 1; }
        public byte[] GetBlob() { return rawBlob; }
        public byte[] EchoBlob(byte[] value) { return value; }

        public byte[] CreateBlob(int size)
        {
            var blob = new byte[size];
            for (int i = 0; i < size; i++)
                blob[i] = (byte)(i % 256);
            return blob;
        }
    }
}
