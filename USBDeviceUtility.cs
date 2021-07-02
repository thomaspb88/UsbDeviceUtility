using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace UsbDeviceUtility
{
    public class USBDeviceUtility
    {
        private USBProperty _USBWatcherOptions;
        private string _watcherProperty;
        private string _win32Provider;
        private ManagementEventWatcher insertWatcher;
        private ManagementObjectSearcher managementObjectSearcher;
        private ManagementEventWatcher removeWatcher;
        private readonly Dictionary<string, string> usbdeviceDictionary = new Dictionary<string, string>();

        public event EventHandler<USBDeviceConnectionChangedEventArgs> DeviceConnectionChanged;
        public void ConfigureProvider(string win32Provider)
        {
            if (string.IsNullOrWhiteSpace(win32Provider)) throw new ArgumentNullException();
            _win32Provider = win32Provider;
        }

        public List<USBDevice> GetAll()
        {
            this.managementObjectSearcher = new ManagementObjectSearcher($"SELECT * FROM {_win32Provider}");
            var result = managementObjectSearcher.Get();

            var usbDevices = new List<USBDevice>();

            foreach (var obj in result)
            {
                usbDevices.Add(new USBDevice()
                {
                    DeviceID = obj.Properties["DeviceID"].Value.ToString(),
                    Name = obj.Properties["Name"].Value.ToString()
                });
            }
            return usbDevices;
        }

        public USBDevice GetByDeviceID(string DeviceID)
        {
            if (string.IsNullOrWhiteSpace(DeviceID)) throw new ArgumentNullException("DeviceID is null or whitespace");

            var deviceQuery = $"SELECT * FROM {_win32Provider} WHERE DeviceID = '{DeviceID}' LIMIT 1";
            return FindDevice(deviceQuery);
        }

        public USBDevice GetByName(string Name)
        {
            if (string.IsNullOrWhiteSpace(Name)) throw new ArgumentNullException("Name is null or whitespace");

            var deviceQuery = $"SELECT * FROM {_win32Provider} WHERE Name = '{Name}' LIMIT 1";

            return FindDevice(deviceQuery);
        }

        public bool IsConnected()
        {
            switch (_USBWatcherOptions)
            {
                case USBProperty.Name:
                    return usbdeviceDictionary.ContainsValue(_watcherProperty);

                case USBProperty.DeviceID:
                default:
                    return usbdeviceDictionary.ContainsKey(_watcherProperty);
            }
        }

        public void Watch(USBProperty usbProperty, string property)
        {
            if (string.IsNullOrWhiteSpace(_win32Provider)) throw new ArgumentNullException("Unable to watch for device activity without setting the provide");
            if (string.IsNullOrWhiteSpace(property)) throw new ArgumentNullException("Unable to watch for device activity without setting the provide");

            _USBWatcherOptions = usbProperty;
            _watcherProperty = property;

            WqlEventQuery insertQuery = new WqlEventQuery($"SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA '{_win32Provider}'");

            insertWatcher = new ManagementEventWatcher(insertQuery);
            insertWatcher.EventArrived += DeviceInsertedEvent;
            insertWatcher.Start();

            WqlEventQuery removeQuery = new WqlEventQuery($"SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA '{_win32Provider}'");
            removeWatcher = new ManagementEventWatcher(removeQuery);
            removeWatcher.EventArrived += DeviceRemovedEvent;
            removeWatcher.Start();

            foreach (var obj in GetAllDevices())
            {
                usbdeviceDictionary[obj.Properties["DeviceID"].Value.ToString()] = obj.Properties["Name"].Value.ToString();
            }
        }

        protected virtual void OnDeviceConnectionChanged(USBDeviceConnectionChangedEventArgs args)
        {
            EventHandler<USBDeviceConnectionChangedEventArgs> handler = DeviceConnectionChanged;
            handler?.Invoke(this, args);
        }
        #region EventHandlers

        private void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            string result = ExtractDeviceIDFromManagementBaseObject(instance);

            ManagementObjectCollection objs = FindPnPEntityManagementObjectByDeviceID(result);
            foreach (var obj in objs)
            {
                usbdeviceDictionary[obj.Properties["DeviceID"].Value.ToString()] = obj.Properties["Name"].Value.ToString();
                OnDeviceConnectionChanged(CreateUSBConnectionChangedEventArgs(USBState.Connected, obj.Properties["Name"].Value.ToString(), obj.Properties["DeviceID"].Value.ToString()));
            }
        }

        private void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];

            string deviceID = ExtractDeviceIDFromManagementBaseObject(instance).Replace("\\\\", "\\");

            switch (_USBWatcherOptions)
            {
                case USBProperty.DeviceID:
                    FilterDeviceIDRaiseEventIfMatchFound(deviceID);
                    break;

                case USBProperty.Name:
                default:
                    FilterNameRaiseEventIfMatchFound(deviceID);
                    break;
            }
        }

        private void FilterDeviceIDRaiseEventIfMatchFound(string deviceID)
        {
            if (usbdeviceDictionary.ContainsKey(deviceID) && _watcherProperty == usbdeviceDictionary[deviceID])
            {
                usbdeviceDictionary.Remove(deviceID);
                OnDeviceConnectionChanged(CreateUSBConnectionChangedEventArgs(USBState.Disconnected, usbdeviceDictionary[deviceID], deviceID));
            }
        }

        private void FilterNameRaiseEventIfMatchFound(string deviceID)
        {
            if (usbdeviceDictionary.ContainsKey(deviceID) && _watcherProperty == usbdeviceDictionary[deviceID])
            {
                OnDeviceConnectionChanged(CreateUSBConnectionChangedEventArgs(USBState.Disconnected, usbdeviceDictionary[deviceID], deviceID));
                usbdeviceDictionary.Remove(deviceID);
            }
        }
        private ManagementObjectCollection GetAllDevices()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Name IS NOT NUll");
            return searcher.Get();
        }

        #endregion EventHandlers

        #region Helper Classes

        private USBDeviceConnectionChangedEventArgs CreateUSBConnectionChangedEventArgs(USBState state, string name, string deviceID)
        {
            if (name == null || deviceID == null) return null;
            return new USBDeviceConnectionChangedEventArgs()
            {
                USBDevice = new USBDevice()
                {
                    DeviceID = deviceID,
                    Name = name
                },
                USBState = state
            };
        }

        private string ExtractDeviceIDFromManagementBaseObject(ManagementBaseObject instance)
        {
            return instance["Dependent"].ToString().Split(new string[] { "DeviceID=" }, StringSplitOptions.RemoveEmptyEntries)[1].Trim('"');
        }
        private ManagementObjectCollection FindPnPEntityManagementObjectByDeviceID(string deviceID)
        {
            ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher($"SELECT * FROM Win32_PnPEntity WHERE DeviceID = '{deviceID}'");
            return managementObjectSearcher.Get();
        }

        private USBDevice FindDevice(string deviceQuery)
        {
            managementObjectSearcher = new ManagementObjectSearcher(deviceQuery);
            var result = managementObjectSearcher.Get().OfType<ManagementBaseObject>().FirstOrDefault();

            if (result == null) return null;

            return new USBDevice()
            {
                DeviceID = result.Properties["DeviceID"].Value.ToString(),
                Name = result.Properties["Name"].Value.ToString()
            };
        }

        #endregion Helper Classes
    }
}