using System;
using iTunesLib;
using System.Security.Cryptography;
using System.Text;
using System.Linq;
using System.IO;

namespace iTunesNowPlaying
{
    public class ArtworkCache
    {
        public static string CacheFolder => Environment.CurrentDirectory + "\\artwork\\";
        private const string DefaultArtworkName = "artwork";

        public ArtworkCache()
        {
            // キャッシュフォルダを作る
            Directory.CreateDirectory(CacheFolder);
        }

        public string GetArtworkPath(IITTrack track)
        {
            IITArtworkCollection artworks = track.Artwork;
            if (artworks?.Count > 0)
            {
                var enumerator = artworks.GetEnumerator();
                enumerator.MoveNext();
                IITArtwork artwork = (IITArtwork)enumerator.Current;
                string hash = ComputeHash(track.Album, track.Artist);
                if (hash == null)
                {
                    // ハッシュを算出できなかった場合はデフォルトのアートワーク名で
                    // 保存して返す
                    string path = BuildArtworkPath(artwork, DefaultArtworkName);
                    artwork.SaveArtworkToFile(path);
                    return path;
                }
                else
                {
                    string path = BuildArtworkPath(artwork, hash);
                    // キャッシュフォルダにすでにアートワークが存在するか?
                    if (File.Exists(path))
                    {
                        // キャッシュに存在するならすぐパスを返す
                        return path;
                    }
                    else
                    {
                        // キャッシュに存在しないならば保存して返す
                        artwork.SaveArtworkToFile(path);
                        return path;
                    }
                }
            }
            else
            {
                return null;
            }
        }

        private string BuildArtworkPath(IITArtwork artwork, string fileName)
        {
            return ArtworkCache.CacheFolder + fileName + GetArtworkExtension(artwork.Format);
        }

        /// <summary>
        /// Convert ITArtworkformat to extension as string
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
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

        private string ComputeHash(string album, string artist)
        {
            if (!string.IsNullOrEmpty(album) && !string.IsNullOrEmpty(artist))
            {
                SHA256 sha256 = new SHA256CryptoServiceProvider();
                var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(album + artist))
                    .Take(8)
                    .Select(x => string.Format("{0:X2}", x))
                    .Aggregate((a, b) => a + b);
                return hash;
            }
            else
            {
                return null;
            }
        }

        public void CleanCache()
        {
            foreach (var filePath in Directory.GetFiles(CacheFolder))
            {
                File.SetAttributes(filePath, FileAttributes.Normal);
                File.Delete(filePath);
            }
        }
    }
}
