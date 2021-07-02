using System;

namespace UsbDeviceUtility
{
    public class USBDeviceConnectionChangedEventArgs : EventArgs
    {
        public USBDevice USBDevice { get; set; }
        public USBState USBState { get; set; }
    }
}