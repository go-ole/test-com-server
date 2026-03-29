using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000E"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestConnectionPoint
    {
        int QueryValue();
        void FireEvent();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000F"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICOMTestConnectionPointEvents
    {
        [DispId(1)]
        void OnEvent(int value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000010"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.ConnectionPoint"),
        ComSourceInterfaces(typeof(ICOMTestConnectionPointEvents))]
    public class COMTestConnectionPoint : ICOMTestConnectionPoint
    {
        public delegate void OnEventHandler(int value);
        public event OnEventHandler OnEvent;

        public COMTestConnectionPoint() { }

        public int QueryValue() { return 999; }

        public void FireEvent()
        {
            OnEvent?.Invoke(999);
        }
    }
}
