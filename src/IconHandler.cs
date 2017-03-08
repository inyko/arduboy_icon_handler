using SharpShell.Attributes;
using SharpShell.SharpIconHandler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace Arduboy
{
    #region Information

    [DataContract]
    internal sealed class Info
    {
        [Serializable]
        public class IconDict : ISerializable
        {
            private Dictionary<string, string> dict = new Dictionary<string, string>();

            public string this[string key] { get { return dict[key]; } }

            public IconDict()
            {
            }

            protected IconDict(SerializationInfo info, StreamingContext context)
            {
                foreach (var entry in info)
                {
                    string str = entry.Value as string;
                    dict.Add(entry.Name, str);
                }
            }

            public bool ContainsKey(string key)
            {
                return dict.ContainsKey(key);
            }

            public void GetObjectData(SerializationInfo info, StreamingContext context)
            {
                foreach (string key in dict.Keys)
                {
                    info.AddValue(key, dict[key]);
                }
            }
        }

        [DataMember(Name = "icons", IsRequired = false)]
        public IconDict Icons { get; set; }

        public string Windows
        {
            get
            {
                if (Icons == null)
                    return null;
                if (!Icons.ContainsKey("windows"))
                    return null;

                return Icons["windows"];
            }
        }

        public string Default
        {
            get
            {
                if (Icons == null)
                    return null;
                if (!Icons.ContainsKey("default"))
                    return null;

                return Icons["default"];
            }
        }
    }

    #endregion

    #region Icon converter

    internal class PngIconConverter
    {
        /* Input image with width = height is suggested to get the best result */
        /* png support in icon was introduced in Windows Vista. */
        public static bool Convert(Stream input_stream, Stream output_stream, int size, bool keep_aspect_ratio = false)
        {
            Bitmap input_bit = (Bitmap)Bitmap.FromStream(input_stream);
            if (input_bit != null)
            {
                int width, height;
                if (keep_aspect_ratio)
                {
                    width = size;
                    height = input_bit.Height / input_bit.Width * size;
                }
                else
                {
                    width = height = size;
                }
                Bitmap new_bit = new Bitmap(input_bit, new Size(width, height));
                if (new_bit != null)
                {
                    // save the resized png into a memory stream for future use
                    MemoryStream mem_data = new MemoryStream();
                    new_bit.Save(mem_data, System.Drawing.Imaging.ImageFormat.Png);

                    BinaryWriter icon_writer = new BinaryWriter(output_stream);
                    if (output_stream != null && icon_writer != null)
                    {
                        // 0-1 reserved, 0
                        icon_writer.Write((byte)0);
                        icon_writer.Write((byte)0);

                        // 2-3 image type, 1 = icon, 2 = cursor
                        icon_writer.Write((short)1);

                        // 4-5 number of images
                        icon_writer.Write((short)1);

                        // image entry 1
                        // 0 image width
                        icon_writer.Write((byte)width);
                        // 1 image height
                        icon_writer.Write((byte)height);

                        // 2 number of colors
                        icon_writer.Write((byte)0);

                        // 3 reserved
                        icon_writer.Write((byte)0);

                        // 4-5 color planes
                        icon_writer.Write((short)0);

                        // 6-7 bits per pixel
                        icon_writer.Write((short)32);

                        // 8-11 size of image data
                        icon_writer.Write((int)mem_data.Length);

                        // 12-15 offset of image data
                        icon_writer.Write((int)(6 + 16));

                        // write image data
                        // png data must contain the whole png data file
                        icon_writer.Write(mem_data.ToArray());

                        icon_writer.Flush();

                        return true;
                    }
                }

                return false;
            }

            return false;
        }

        public static Icon Convert(Stream inputStreem, int size)
        {
            using (MemoryStream memStream = new MemoryStream())
            {
                if (!Convert(inputStreem, memStream, size))
                    return null;

                memStream.Seek(0, SeekOrigin.Begin);
                Icon result = new Icon(memStream);

                return result;
            }
        }
    }

    #endregion

    [ComVisible(true)]
    [COMServerAssocation(AssociationType.ClassOfExtension, ".arduboy")]
    public class IconHandler : SharpIconHandler
    {
        protected override Icon GetIcon(bool smallIcon, uint iconSize)
        {
            Icon icon = ArduboyIconHandler.Properties.Resources.Arduino;
            try
            {
                icon = ExtractIcon(SelectedItemPath);
            }
            catch
            {
            }

            return GetIconSpecificSize(icon, new Size((int)iconSize, (int)iconSize));
        }

        private Icon ExtractIcon(string path)
        {
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry infoEntry = archive.GetEntry("info.json");
                if (infoEntry == null) return ArduboyIconHandler.Properties.Resources.Arduino;
                Info info = null;
                using (Stream infoStream = infoEntry.Open())
                {
                    DataContractJsonSerializer json = new DataContractJsonSerializer(typeof(Info));
                    info = json.ReadObject(infoStream) as Info;
                }
                if (info == null) return ArduboyIconHandler.Properties.Resources.Arduino;
                Icon icon = LoadIcon(archive, info.Windows);
                if (icon != null) return icon;
                icon = LoadIcon(archive, info.Default);
                if (icon != null) return icon;
                icon = LoadIcon(archive, "icon.ico");
                if (icon != null) return icon;
                icon = LoadIcon(archive, "icon.png");
                if (icon != null) return icon;
            }

            return ArduboyIconHandler.Properties.Resources.Arduino;
        }

        private Icon LoadIcon(ZipArchive archive, string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            ZipArchiveEntry iconEntry = archive.GetEntry(path);
            if (iconEntry == null) return null;
            using (Stream iconStream = iconEntry.Open())
            {
                using (MemoryStream mem = new MemoryStream())
                {
                    iconStream.CopyTo(mem);
                    mem.Seek(0, SeekOrigin.Begin);

                    using (Bitmap img = Bitmap.FromStream(mem) as Bitmap)
                    {
                        IntPtr Hicon = img.GetHicon();
                        Icon icon1 = Icon.FromHandle(Hicon);
                        Icon icon2 = null;

                        try
                        {
                            mem.Seek(0, SeekOrigin.Begin);
                            icon2 = PngIconConverter.Convert(mem, img.Width);
                        }
                        catch
                        {
                        }

                        return icon2 ?? icon1;
                    }
                }
            }
        }
    }
}
