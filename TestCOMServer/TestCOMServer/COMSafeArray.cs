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
    }

    [ComVisible(true),
        Guid("902A4816-13F3-4DA1-B712-B80B149D7F35"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ComSourceInterfaces(typeof(ICOMSafeArray))]
    class COMSafeArray : ICOMSafeArray
    {
    }
}
