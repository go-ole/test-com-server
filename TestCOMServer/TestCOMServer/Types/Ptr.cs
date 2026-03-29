using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000002C"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICOMTestPtr
    {
        IntPtr GetPtr();
        IntPtr EchoPtr(IntPtr value);
        int PutPtr(IntPtr value);
        IntPtr GetStoredPtr();
        IntPtr AllocatePtr(int value);
        int ReadPtr(IntPtr ptr);
        void FreePtr(IntPtr ptr);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000002D"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Ptr"),
        ComSourceInterfaces(typeof(ICOMTestPtr))]
    public class COMTestPtr : ICOMTestPtr
    {
        protected IntPtr storedPtr;

        public COMTestPtr()
        {
            storedPtr = IntPtr.Zero;
        }

        public IntPtr GetPtr()
        {
            return new IntPtr(0x12345678);
        }

        public IntPtr EchoPtr(IntPtr value) { return value; }
        public int PutPtr(IntPtr value) { storedPtr = value; return 1; }
        public IntPtr GetStoredPtr() { return storedPtr; }

        public IntPtr AllocatePtr(int value)
        {
            IntPtr ptr = Marshal.AllocCoTaskMem(sizeof(int));
            Marshal.WriteInt32(ptr, value);
            return ptr;
        }

        public int ReadPtr(IntPtr ptr)
        {
            return Marshal.ReadInt32(ptr);
        }

        public void FreePtr(IntPtr ptr)
        {
            Marshal.FreeCoTaskMem(ptr);
        }
    }
}
