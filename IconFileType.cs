using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using PaintDotNet;
using PaintDotNet.IndirectUI;
using PaintDotNet.PropertySystem;
using PaintDotNet.Rendering;

namespace IconCreatorFileType
{
    public class FileFormatException : Exception
    {
        public FileFormatException() : base() { }
        public FileFormatException(string message) : base(message) { }
        public FileFormatException(string message, Exception inner) : base(message, inner) { }
    }

    public enum ImageSize
    {
        Icon_Auto,
        Icon_16x16,
        Icon_32x32,
        Icon_48x48,
        Icon_64x64,
        Icon_128x128,
        Icon_256x256,
        Icon_All
    }

    [PluginSupportInfo(typeof(PluginSupportInfo))]
    public sealed class IconFileType : PropertyBasedFileType
    {

        public const string ImageSizeString = "Image Size";
        public const string LessAndEqualSizesString = "Stack Icons";
        public const string SourceCodeString = "Source Code";
        public const string WebSiteLinkValue = @"https://github.com/KyleTTucker/Paint.net-Icon-Creator";

        public IconFileType() : base("Icon", new FileTypeOptions()
        {
            LoadExtensions = new string[] { ".ico" },
            SaveExtensions = new string[] { ".ico" },
            SupportsCancellation = true,
            SupportsLayers = false
        })
        {
        }

        public override PropertyCollection OnCreateSavePropertyCollection()
        {
            List<Property> props =
            [
                StaticListChoiceProperty.CreateForEnum(ImageSizeString, ImageSize.Icon_256x256),
                new BooleanProperty(LessAndEqualSizesString), // Property.Create(typeof(bool), LessAndEqualSizesString, true),
                new UriProperty(SourceCodeString, new Uri(WebSiteLinkValue))
            ];

            return new PropertyCollection(props);
        }

        public override ControlInfo OnCreateSaveConfigUI(PropertyCollection props)
        {
            ControlInfo configUI = CreateDefaultSaveConfigUI(props);
            configUI.SetPropertyControlValue(LessAndEqualSizesString, ControlInfoPropertyNames.DisplayName, string.Empty); // Hide the label
            configUI.SetPropertyControlValue(LessAndEqualSizesString, ControlInfoPropertyNames.Description, "Add all the smaller sizes to the icons"); // Set the description

            return configUI;
        }

        protected override Document OnLoad(Stream input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));
            if (!input.CanSeek)
            {
                var ms = new MemoryStream();
                input.CopyTo(ms);
                ms.Position = 0;
                input = ms;
            }

            try
            {
                var entries = ParseIcoEntries(input);
                if (entries == null || entries.Count == 0)
                    throw new FileFormatException("No icon entries found in ICO stream.");

                var extracted = new List<ExtractedIcon>();
                foreach (var e in entries)
                {
                    try
                    {
                        var bmp = BuildBitmapForEntry(input, e);
                        if (bmp != null)
                        {
                            extracted.Add(new ExtractedIcon
                            {
                                Width = e.Width,
                                Height = e.Height,
                                BitCount = e.BitCount,
                                Bitmap = bmp
                            });
                        }
                    }
                    catch
                    {
                        // ignore individual malformed entries
                    }
                }

                if (extracted.Count == 0)
                    throw new FileFormatException("Could not decode any icon images from ICO.");

                // Select largest icon (area), tie-breaker by bit depth
                var best = extracted
                    .OrderByDescending(x => (long)x.Width * x.Height)
                    .ThenByDescending(x => x.BitCount)
                    .First();

                // Convert Bitmap to Paint.NET Document
                Document doc = null;
                Bitmap bitmap = best.Bitmap;
                if (bitmap.Width > 0 && bitmap.Height > 0)
                {
                    doc = new Document(bitmap.Width, bitmap.Height);
                    BitmapLayer layer = Layer.CreateBackgroundLayer(bitmap.Width, bitmap.Height);
                    Surface surface = layer.Surface;

                    for (int y = 0; y < surface.Height; y++)
                    {
                        for (int x = 0; x < surface.Width; x++)
                        {
                            System.Drawing.Color c = bitmap.GetPixel(x, y);
                            surface[x, y] = ColorBgra.FromBgra(c.B, c.G, c.R, c.A);
                        }
                    }

                    doc.Layers.Add(layer);
                }

                // Dispose other bitmaps
                foreach (var ex in extracted)
                {
                    if (!ReferenceEquals(ex.Bitmap, bitmap))
                        ex.Bitmap.Dispose();
                }

                return doc;
            }
            catch (Exception ex)
            {
                throw new FileFormatException("Failed to load ICO: " + ex.Message, ex);
            }
        }

        #region ICO helpers

        private class IcoDirEntry
        {
            public int Width;
            public int Height;
            public int ColorCount;
            public int Reserved;
            public ushort Planes;
            public ushort BitCount;
            public uint BytesInRes;
            public uint ImageOffset;
        }

        private class ExtractedIcon
        {
            public int Width;
            public int Height;
            public int BitCount;
            public Bitmap Bitmap;
        }

        private List<IcoDirEntry> ParseIcoEntries(Stream s)
        {
            long origPos = s.Position;
            try
            {
                using (var br = new BinaryReader(s, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    s.Position = 0;
                    ushort reserved = br.ReadUInt16();
                    ushort type = br.ReadUInt16();
                    ushort count = br.ReadUInt16();

                    var entries = new List<IcoDirEntry>();
                    for (int i = 0; i < count; i++)
                    {
                        byte w = br.ReadByte();
                        byte h = br.ReadByte();
                        byte colorCount = br.ReadByte();
                        byte reservedByte = br.ReadByte();
                        ushort planes = br.ReadUInt16();
                        ushort bitCount = br.ReadUInt16();
                        uint bytesInRes = br.ReadUInt32();
                        uint imageOffset = br.ReadUInt32();

                        int width = (w == 0) ? 256 : w;
                        int height = (h == 0) ? 256 : h;

                        if (imageOffset == 0 || bytesInRes == 0)
                            continue;

                        entries.Add(new IcoDirEntry
                        {
                            Width = width,
                            Height = height,
                            ColorCount = colorCount,
                            Reserved = reservedByte,
                            Planes = planes,
                            BitCount = bitCount,
                            BytesInRes = bytesInRes,
                            ImageOffset = imageOffset
                        });
                    }

                    return entries;
                }
            }
            finally
            {
                s.Position = origPos;
            }
        }

        private Bitmap BuildBitmapForEntry(Stream fullIcoStream, IcoDirEntry entry)
        {
            byte[] imageBytes = new byte[entry.BytesInRes];
            fullIcoStream.Position = entry.ImageOffset;
            int read = 0;
            while (read < imageBytes.Length)
            {
                int r = fullIcoStream.Read(imageBytes, read, imageBytes.Length - read);
                if (r == 0) break;
                read += r;
            }
            if (read != imageBytes.Length)
                throw new EndOfStreamException("Unexpected end of ICO when reading image data");

            using (var ms = new MemoryStream())
            using (var bw = new BinaryWriter(ms))
            {
                bw.Write((ushort)0); // reserved
                bw.Write((ushort)1); // type
                bw.Write((ushort)1); // count

                byte w = (byte)(entry.Width == 256 ? 0 : entry.Width);
                byte h = (byte)(entry.Height == 256 ? 0 : entry.Height);
                bw.Write(w);
                bw.Write(h);
                bw.Write((byte)entry.ColorCount);
                bw.Write((byte)entry.Reserved);
                bw.Write((ushort)entry.Planes);
                bw.Write((ushort)entry.BitCount);
                bw.Write((uint)imageBytes.Length);
                bw.Write((uint)(6 + 16)); // offset to image data
                bw.Write(imageBytes);

                bw.Flush();
                ms.Position = 0;

                using (var icon = new Icon(ms))
                {
                    return icon.ToBitmap();
                }
            }
        }

        #endregion

        protected override void OnSaveT(Document input, Stream output, PropertyBasedSaveConfigToken token, Surface scratchSurface, ProgressEventHandler progressCallback)
        {
            scratchSurface.Clear();
            input.CreateRenderer().Render(scratchSurface);


            // Render the source image into a square (max(width, height)), centered, to avoid stretching
            int maxDim = Math.Max(scratchSurface.Width, scratchSurface.Height);
            Bitmap ApplyPixels = new(maxDim, maxDim);
            using (Graphics g = Graphics.FromImage(ApplyPixels))
            {
                g.Clear(Color.Transparent);
                int offsetX = (maxDim - scratchSurface.Width) / 2;
                int offsetY = (maxDim - scratchSurface.Height) / 2;
                // Copy the pixels from scratchSurface into the centered square
                for (int i = 0; i < scratchSurface.Width; i++)
                {
                    for (int j = 0; j < scratchSurface.Height; j++)
                    {
                        ApplyPixels.SetPixel(i + offsetX, j + offsetY, scratchSurface[i, j]);
                    }
                }
            }

            //Resize image
            ImageSize bitDepth = (ImageSize)token.GetProperty(ImageSizeString).Value;
            bool stackIcons = (bool)token.GetProperty(LessAndEqualSizesString).Value;
            if (bitDepth == ImageSize.Icon_All || stackIcons)
            {
                // Refactored: efficient generation of a multi-resolution icon
                int[] sizes = new int[] { 16, 32, 48, 64, 128, 256 };

                // If stacking icons, filter sizes based on selected size
                if (stackIcons && bitDepth != ImageSize.Icon_All)
                {
                    // Only include sizes less than or equal to the selected size
                    int maxSize = bitDepth switch
                    {
                        ImageSize.Icon_16x16 => 16,
                        ImageSize.Icon_32x32 => 32,
                        ImageSize.Icon_48x48 => 48,
                        ImageSize.Icon_64x64 => 64,
                        ImageSize.Icon_128x128 => 128,
                        ImageSize.Icon_256x256 => 256,
                        _ => 32,
                    };
                    sizes = Array.FindAll(sizes, sz => sz <= maxSize);
                }

                var pngImages = new List<byte[]>(sizes.Length);
                using (var _iconWriter = new BinaryWriter(output, System.Text.Encoding.Default, true))
                {
                    // Generate the bitmaps and PNGs in memory
                    foreach (int sz in sizes)
                    {
                        using (var bmp = new Bitmap(ApplyPixels, sz, sz))
                        using (var ms = new MemoryStream())
                        {
                            bmp.Save(ms, ImageFormat.Png);
                            pngImages.Add(ms.ToArray());
                        }
                    }

                    // ICO header
                    _iconWriter.Write((short)0); // reserved
                    _iconWriter.Write((short)1); // type: icon
                    _iconWriter.Write((short)sizes.Length); // number of images

                    int offset = 6 + (16 * sizes.Length); // 6 bytes header + 16 bytes per entry
                    for (int i = 0; i < sizes.Length; i++)
                    {
                        int sz = sizes[i];
                        byte width = (byte)(sz == 256 ? 0 : sz); // 0 = 256px
                        byte height = (byte)(sz == 256 ? 0 : sz); // 0 = 256px
                        _iconWriter.Write(width); // width
                        _iconWriter.Write(height); // height
                        _iconWriter.Write((byte)0); // colors
                        _iconWriter.Write((byte)0); // reserved
                        _iconWriter.Write((short)0); // color planes
                        _iconWriter.Write((short)32); // bits per pixel
                        _iconWriter.Write(pngImages[i].Length); // image size
                        _iconWriter.Write(offset); // offset
                        offset += pngImages[i].Length;
                    }
                    // PNG data
                    for (int i = 0; i < pngImages.Count; i++)
                    {
                        _iconWriter.Write(pngImages[i]);
                    }
                    _iconWriter.Flush();
                }
                return;
            }
            switch (bitDepth)
            {
                case ImageSize.Icon_Auto:
                case ImageSize.Icon_32x32:
                    ApplyPixels = new Bitmap(ApplyPixels, 32, 32);
                    break;
                case ImageSize.Icon_16x16:
                    ApplyPixels = new Bitmap(ApplyPixels, 16, 16);
                    break;
                case ImageSize.Icon_48x48:
                    ApplyPixels = new Bitmap(ApplyPixels, 48, 48);
                    break;
                case ImageSize.Icon_64x64:
                    ApplyPixels = new Bitmap(ApplyPixels, 64, 64);
                    break;
                case ImageSize.Icon_128x128:
                    ApplyPixels = new Bitmap(ApplyPixels, 128, 128);
                    break;
                case ImageSize.Icon_256x256:
                    ApplyPixels = new Bitmap(ApplyPixels, 256, 256);
                    break;
                case ImageSize.Icon_All:

                default:
                    ApplyPixels = new Bitmap(ApplyPixels, 32, 32);
                    break;
            }

            BinaryWriter iconWriter = new(output);

            //Check for any null streams
            if (iconWriter == null || output == null)
                return;

            MemoryStream memoryStream = new();
            ApplyPixels.Save(memoryStream, ImageFormat.Png);

            //https://fileformats.fandom.com/wiki/Icon
            // Icon file format

            // 0-1 reserved, 0
            iconWriter.Write((short)0);

            // 2-3 image type, 1 = icon, 2 = cursor
            iconWriter.Write((short)1);

            // 4-5 number of images
            iconWriter.Write((short)1);

            // 0 image width
            iconWriter.Write((byte)ApplyPixels.Width);

            // 1 image height
            iconWriter.Write((byte)ApplyPixels.Height);

            // 2 number of colors
            iconWriter.Write((byte)0);

            // 3 reserved
            iconWriter.Write((byte)0);

            // 4-5 color planes
            iconWriter.Write((short)0);

            // 6-7 bits per pixel
            iconWriter.Write((short)32);

            // 8-11 size of image data
            iconWriter.Write((int)memoryStream.Length);

            // 12-15 offset of image data
            iconWriter.Write((int)22);

            iconWriter.Write(memoryStream.ToArray());
            memoryStream.Close();

            iconWriter.Flush();
        }
    }
}
