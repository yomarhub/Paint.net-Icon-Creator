using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using PaintDotNet;
using PaintDotNet.PropertySystem;
using PaintDotNet.Rendering;

namespace IconCreatorFileType
{
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
                StaticListChoiceProperty.CreateForEnum(ImageSizeString, ImageSize.Icon_Auto),
                new UriProperty(SourceCodeString, new Uri(WebSiteLinkValue))
            ];

            return new PropertyCollection(props);
        }

        protected override Document OnLoad(Stream input)
        {
            Icon NewIcon = new(input);
            Bitmap bitmapOfIcon = NewIcon.ToBitmap();

            Document document = null;

            if (bitmapOfIcon.Width > 0 && bitmapOfIcon.Height > 0)
            {
                document = new Document(bitmapOfIcon.Width, bitmapOfIcon.Height);

                BitmapLayer layer = Layer.CreateBackgroundLayer(bitmapOfIcon.Width, bitmapOfIcon.Height);

                Surface surface = layer.Surface;

                for (int y = 0; y < surface.Height; y++)
                {
                    for (int x = 0; x < surface.Width; x++)
                    {
                        surface[x, y] = bitmapOfIcon.GetPixel(x, y);
                    }
                }

                document.Layers.Add(layer);
            }

            return document;
        }

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
                    // Refactored: efficient generation of a multi-resolution icon
                    int[] sizes = new int[] { 16, 32, 48, 64, 128, 256 };
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
