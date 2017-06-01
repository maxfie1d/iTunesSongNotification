using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTunesLib;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using iTunesPlugin.ShellHelpers;
using Microsoft.Toolkit.Uwp.Notifications;
using Windows.UI.Notifications;
using Windows.Data.Xml.Dom;

namespace iTunesPlugin
{
    class Program
    {
        static iTunesApp app;
        static void Main(string[] args)
        {
            // COM初期化
            app = new iTunesApp();
            // イベントハンドラを登録
            app.OnAboutToPromptUserToQuitEvent += new _IiTunesEvents_OnAboutToPromptUserToQuitEventEventHandler(oniTunesQuitting);
            app.OnQuittingEvent += new _IiTunesEvents_OnQuittingEventEventHandler(oniTunesQuitting);
            app.OnPlayerPlayEvent += new _IiTunesEvents_OnPlayerPlayEventEventHandler(app_OnPlayerPlayEvent);

            NotificationActivator.Initialize();
            Console.ReadKey();

            Dispose();
            NotificationActivator.Uninitialize();
        }

        private static void app_OnPlayerPlayEvent(object iTrack)
        {
            var track = (IITTrack)iTrack;
            Console.WriteLine(track.Name);
            ShowSongNotification(track.Name, track.Artist, track.Album, GetArtwork(track));
        }

        private static string GetArtwork(IITTrack track)
        {
            IITArtworkCollection artwork = track.Artwork;
            if (artwork?.Count > 0)
            {
                foreach (IITArtwork a in artwork)
                {
                    string extension = GetArtworkExtension(a.Format);
                    if (extension != null)
                    {
                        string path = Environment.CurrentDirectory + "\\artwork" + extension;
                        a.SaveArtworkToFile(path);
                        return path;
                    }
                    else
                    {
                        return null;
                    }
                }
                return null;
            }
            else
            {
                return null;
            }
        }

        private static string GetArtworkExtension(ITArtworkFormat format)
        {
            switch (format)
            {
                case ITArtworkFormat.ITArtworkFormatUnknown: return null;
                case ITArtworkFormat.ITArtworkFormatJPEG: return ".jpg";
                case ITArtworkFormat.ITArtworkFormatPNG: return ".png";
                case ITArtworkFormat.ITArtworkFormatBMP: return ".bmp";
                default:
                    return null;
            }
        }

        private static void oniTunesQuitting()
        {
            Dispose();
        }

        private static void Dispose()
        {
            if (app != null)
            {
                app.OnAboutToPromptUserToQuitEvent -= oniTunesQuitting;
                app.OnQuittingEvent -= oniTunesQuitting;
                app.OnPlayerPlayEvent -= app_OnPlayerPlayEvent;

                // COMを解放
                Marshal.ReleaseComObject(app);
                app = null;
            }
        }

        private static void RegisterAppForNotificationSupport()
        {
            String shortcutPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Windows\\Start Menu\\Programs\\Nitro Desktop Toasts Sample CS.lnk";
            if (!File.Exists(shortcutPath))
            {
                // Find the path to the current executable
                String exePath = Process.GetCurrentProcess().MainModule.FileName;
                InstallShortcut(shortcutPath, exePath);
                RegisterComServer(exePath);
            }
        }

        private static void InstallShortcut(String shortcutPath, String exePath)
        {
            IShellLinkW newShortcut = (IShellLinkW)new CShellLink();

            // Create a shortcut to the exe
            newShortcut.SetPath(exePath);

            // Open the shortcut property store, set the AppUserModelId property
            IPropertyStore newShortcutProperties = (IPropertyStore)newShortcut;

            PropVariantHelper varAppId = new PropVariantHelper();
            varAppId.SetValue(APP_ID);
            newShortcutProperties.SetValue(PROPERTYKEY.AppUserModel_ID, varAppId.Propvariant);

            PropVariantHelper varToastId = new PropVariantHelper();
            varToastId.VarType = VarEnum.VT_CLSID;
            varToastId.SetValue(typeof(NotificationActivator).GUID);

            newShortcutProperties.SetValue(PROPERTYKEY.AppUserModel_ToastActivatorCLSID, varToastId.Propvariant);

            // Commit the shortcut to disk
            ShellHelpers.IPersistFile newShortcutSave = (ShellHelpers.IPersistFile)newShortcut;

            newShortcutSave.Save(shortcutPath, true);
        }

        private static void RegisterComServer(String exePath)
        {
            // We register the app process itself to start up when the notification is activated, but
            // other options like launching a background process instead that then decides to launch
            // the UI as needed.
            string regString = String.Format("SOFTWARE\\Classes\\CLSID\\{{{0}}}\\LocalServer32", typeof(NotificationActivator).GUID);
            var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regString);
            key.SetValue(null, exePath);
        }

        private const String APP_ID = "Now playing";

        private static void ShowSongNotification(string title, string artist, string album, string artwork)
        {
            string artistAndAlbum = $"{artist} - {album}";
            ShowNotification(title, artistAndAlbum, artwork);
        }

        private static ToastGenericAppLogo f(string s)
        {
            if (s != null)
            {
                return new ToastGenericAppLogo()
                {
                    Source = s
                };
            }
            else
            {
                return null;
            }
        }

        // Create and show the toast.
        // See the "Toasts" sample for more detail on what can be done with toasts
        private static void ShowNotification(string title, string content, string image)
        {
            ToastContent toastContent = new ToastContent()
            {
                Visual = new ToastVisual()
                {
                    BindingGeneric = new ToastBindingGeneric()
                    {
                        Children =
                        {
                            new AdaptiveText()
                            {
                                Text = title
                            },

                            new AdaptiveText()
                            {
                                Text = content
                            }
                        },
                        AppLogoOverride = f(image),
                    }
                },
                Audio = new ToastAudio
                {
                    Silent = true,
                    Loop = false
                }
            };

            // Create XML from the toast content
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(toastContent.GetContent());

            // Create the toast and attach event listeners
            ToastNotification toast = new ToastNotification(xml);

            toast.Failed += (s, e) =>
            {
                Console.WriteLine("Toast Failed");
            };

            // Show the toast. Be sure to specify the AppUserModelId on your application's shortcut!
            ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
        }
    }
}
