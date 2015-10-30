using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A7199294-BD75-4D87-A308-5696C2ED64E8"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArray
    {
        sbyte[] EchoInt8Array(sbyte[] value);
        byte[] EchoUInt8(byte[] value);
        short[] EchoInt16(short[] value);
        ushort[] EchoUInt16(ushort[] value);
        int[] EchoInt32(int[] value);
        uint[] EchoUInt32(uint[] value);
        long[] EchoInt64(long[] value);
        ulong[] EchoUInt64(ulong[] value);
        float[] EchoFloat32(float[] value);
        double[] EchoFloat64(double[] value);
        string[] EchoString(string[] value);
    }

    [ComVisible(true),
        Guid("902A4816-13F3-4DA1-B712-B80B149D7F35"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ComSourceInterfaces(typeof(ICOMSafeArray))]
    class COMSafeArray : ICOMSafeArray
    {
        public sbyte[] EchoInt8Array(sbyte[] value)
        {
            return value;
        }

        public byte[] EchoUInt8(byte[] value)
        {
            return value;
        }

        public short[] EchoInt16(short[] value)
        {
            return value;
        }

        public ushort[] EchoUInt16(ushort[] value)
        {
            return value;
        }

        public int[] EchoInt32(int[] value)
        {
            return value;
        }

        public uint[] EchoUInt32(uint[] value)
        {
            return value;
        }

        public long[] EchoInt64(long[] value)
        {
            return value;
        }

        public ulong[] EchoUInt64(ulong[] value)
        {
            return value;
        }

        public float[] EchoFloat32(float[] value)
        {
            return value;
        }

        public double[] EchoFloat64(double[] value)
        {
            return value;
        }

        public string[] EchoString(string[] value)
        {
            return value;
        }
    }
}
