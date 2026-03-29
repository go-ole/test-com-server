using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("0000000C-0000-0000-C000-000000000046"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStreamCompat
    {
        void Read([Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                  int cb,
                  IntPtr pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                   int cb,
                   IntPtr pcbWritten);
        void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition);
        void SetSize(long libNewSize);
        void CopyTo(IStreamCompat pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten);
        void Commit(int grfCommitFlags);
        void Revert();
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out STATSTG pstatstg, int grfStatFlag);
        void Clone(out IStreamCompat ppstm);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000029"),
        ClassInterface(ClassInterfaceType.None)]
    public class ManagedStream : IStreamCompat
    {
        private MemoryStream _stream = new MemoryStream();

        public void Read(byte[] pv, int cb, IntPtr pcbRead)
        {
            int bytesRead = _stream.Read(pv, 0, cb);
            if (pcbRead != IntPtr.Zero)
                Marshal.WriteInt32(pcbRead, bytesRead);
        }

        public void Write(byte[] pv, int cb, IntPtr pcbWritten)
        {
            _stream.Write(pv, 0, cb);
            if (pcbWritten != IntPtr.Zero)
                Marshal.WriteInt32(pcbWritten, cb);
        }

        public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
        {
            long pos = _stream.Seek(dlibMove, (SeekOrigin)dwOrigin);
            if (plibNewPosition != IntPtr.Zero)
                Marshal.WriteInt64(plibNewPosition, pos);
        }

        public void SetSize(long libNewSize) { _stream.SetLength(libNewSize); }
        public void CopyTo(IStreamCompat pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten) { }
        public void Commit(int grfCommitFlags) { _stream.Flush(); }
        public void Revert() { }
        public void LockRegion(long libOffset, long cb, int dwLockType) { }
        public void UnlockRegion(long libOffset, long cb, int dwLockType) { }

        public void Stat(out STATSTG pstatstg, int grfStatFlag)
        {
            pstatstg = new STATSTG();
            pstatstg.cbSize = _stream.Length;
            pstatstg.type = 2; // STGTY_STREAM
        }

        public void Clone(out IStreamCompat ppstm) { ppstm = null; }
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000027"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICOMTestStream
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object GetStream();

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object EchoStream([MarshalAs(UnmanagedType.IUnknown)] object stream);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000028"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Stream"),
        ComSourceInterfaces(typeof(ICOMTestStream))]
    public class COMTestStream : ICOMTestStream
    {
        public COMTestStream() { }

        public object GetStream()
        {
            return new ManagedStream();
        }

        public object EchoStream(object stream)
        {
            return stream;
        }
    }
}
