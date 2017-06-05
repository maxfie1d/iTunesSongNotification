using System;
using System.Runtime.InteropServices;

namespace Notification
{
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(INotificationActivationCallback))]
    [Guid("23A5B06E-20BB-4E7E-A0AC-6982ED6A6041"), ComVisible(true)]
    public class NotificationActivator : INotificationActivationCallback
    {
        public void Activate(string appUserModelId, string invokedArgs, NOTIFICATION_USER_INPUT_DATA[] data, uint dataCount)
        {
            Console.WriteLine("The user activated the toast");
        }

        public static void Initialize()
        {
            regService = new RegistrationServices();

            cookie = regService.RegisterTypeForComClients(
                typeof(NotificationActivator),
                RegistrationClassContext.LocalServer,
                RegistrationConnectionType.MultipleUse);
        }
        public static void Uninitialize()
        {
            if (cookie != -1 && regService != null)
            {
                regService.UnregisterTypeForComClients(cookie);
            }
        }

        private static int cookie = -1;
        private static RegistrationServices regService = null;
    }
}
