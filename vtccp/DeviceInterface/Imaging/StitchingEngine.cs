namespace DeviceInterface.Imaging;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

/// <summary>
/// Phase-1 symbol stitching engine.
///
/// Workflow
/// --------
///   1. Caller supplies two JPEG byte arrays (left half, right half of a wide 1D
///      barcode captured in two overlapping passes on the DM475V-LBL).
///   2. <see cref="CorrectSkew"/> rotates each half so bars align vertically.
///      Algorithm: sample the bar-top edge at the left-quarter and right-quarter
///      column positions; compute atan2 of the height difference; apply
///      Graphics.RotateTransform with white fill to preserve image size.
///   3. <see cref="EstimateSeam"/> returns the midpoint of the overlap region —
///      suitable as a default seam before the operator adjusts via slider.
///   4. <see cref="Composite"/> cuts the left image at <paramref name="leftSeam"/>
///      and the right image at <paramref name="rightSeam"/> and concatenates them
///      horizontally into a single Bitmap, returned as JPEG bytes.
///
/// Limitations (Phase 1)
/// ---------------------
///   • Rotation only — no scaling, perspective correction, or sub-pixel alignment.
///   • Manual seam placement — operator sets the cut column via UI slider.
///   • No automatic cross-correlation seam detection (Phase 2).
///   • Requires System.Drawing.Common (already in DeviceInterface.csproj).
///
/// Note: test images (C128 FX label) have not yet been received.  Algorithm
/// parameters may need tuning once real samples are available.
/// See backlog STITCH-1 Phase 2 for automatic seam detection.
/// </summary>
public static class StitchingEngine
{
    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Detects and corrects the skew angle of a single barcode-half image.
    /// Returns corrected JPEG bytes.  If skew detection fails, returns the
    /// original bytes unchanged.
    /// </summary>
    public static byte[] CorrectSkew(byte[] jpegBytes)
    {
        using var bmp = LoadBitmap(jpegBytes);
        double angle  = DetectSkewAngleDegrees(bmp);

        if (Math.Abs(angle) < 0.05)
            return jpegBytes;

        using var rotated = RotateBitmap(bmp, -angle);
        return ToJpeg(rotated);
    }

    /// <summary>
    /// Returns the suggested default seam column for the left image.
    /// Currently: three-quarters along the left image width, giving a 50%
    /// overlap assumption.  Caller should clamp to [0, leftWidth-1].
    /// </summary>
    public static int EstimateSeam(byte[] leftJpeg)
    {
        using var bmp = LoadBitmap(leftJpeg);
        return (int)(bmp.Width * 0.75);
    }

    /// <summary>
    /// Composites left and right halves at the specified seam positions.
    ///
    /// The result is:
    ///   left image columns  [0 .. leftSeam)
    ///   right image columns [rightSeam .. rightImage.Width)
    ///
    /// Both images are scaled to the same height (the taller one wins; the
    /// shorter is padded with white on the top and bottom).
    /// </summary>
    /// <param name="leftJpeg">Skew-corrected left-half JPEG bytes.</param>
    /// <param name="rightJpeg">Skew-corrected right-half JPEG bytes.</param>
    /// <param name="leftSeam">Column in the left image where the cut is made (exclusive).</param>
    /// <param name="rightSeam">Column in the right image where the right half begins (inclusive).</param>
    /// <returns>Composite JPEG bytes ready for IMAGE.LOAD.</returns>
    public static byte[] Composite(
        byte[] leftJpeg,
        byte[] rightJpeg,
        int    leftSeam,
        int    rightSeam)
    {
        using var left  = LoadBitmap(leftJpeg);
        using var right = LoadBitmap(rightJpeg);

        int leftSeamClamped  = Math.Clamp(leftSeam,  0, left.Width  - 1);
        int rightSeamClamped = Math.Clamp(rightSeam, 0, right.Width - 1);

        int leftCols  = leftSeamClamped;
        int rightCols = right.Width - rightSeamClamped;
        int totalW    = leftCols + rightCols;
        int totalH    = Math.Max(left.Height, right.Height);

        using var composite = new Bitmap(totalW, totalH, PixelFormat.Format24bppRgb);
        using var g         = Graphics.FromImage(composite);
        g.Clear(Color.White);

        // Left half: columns 0..leftSeam-1
        var leftSrc  = new Rectangle(0, 0, leftCols, left.Height);
        var leftDst  = new Rectangle(0, (totalH - left.Height) / 2, leftCols, left.Height);
        g.DrawImage(left,  leftDst,  leftSrc,  GraphicsUnit.Pixel);

        // Right half: columns rightSeam..right.Width-1
        var rightSrc = new Rectangle(rightSeamClamped, 0, rightCols, right.Height);
        var rightDst = new Rectangle(leftCols, (totalH - right.Height) / 2, rightCols, right.Height);
        g.DrawImage(right, rightDst, rightSrc, GraphicsUnit.Pixel);

        return ToJpeg(composite);
    }

    // ── Skew detection ────────────────────────────────────────────────────────

    /// <summary>
    /// Estimates the bar-baseline skew angle in degrees by sampling the topmost
    /// dark pixel at the left-quarter and right-quarter x-positions.
    /// Positive angle = right side is lower than left.
    /// </summary>
    private static double DetectSkewAngleDegrees(Bitmap bmp)
    {
        int w = bmp.Width, h = bmp.Height;
        int leftX  = w / 4;
        int rightX = 3 * w / 4;

        int topLeft  = FindFirstDarkRow(bmp, leftX,  threshold: 160);
        int topRight = FindFirstDarkRow(bmp, rightX, threshold: 160);

        if (topLeft < 0 || topRight < 0)
            return 0.0;

        double spanX   = rightX - leftX;
        double spanY   = topRight - topLeft;
        double radians = Math.Atan2(spanY, spanX);
        return radians * (180.0 / Math.PI);
    }

    /// <summary>
    /// Scans the column at <paramref name="x"/> from top to bottom and returns
    /// the y-coordinate of the first pixel whose grayscale value is below
    /// <paramref name="threshold"/> (i.e. dark).  Returns -1 if not found.
    /// </summary>
    private static int FindFirstDarkRow(Bitmap bmp, int x, int threshold)
    {
        x = Math.Clamp(x, 0, bmp.Width - 1);
        for (int y = 0; y < bmp.Height; y++)
        {
            Color c    = bmp.GetPixel(x, y);
            int   gray = (c.R + c.G + c.B) / 3;
            if (gray < threshold) return y;
        }
        return -1;
    }

    // ── Bitmap rotation ───────────────────────────────────────────────────────

    /// <summary>
    /// Rotates <paramref name="src"/> by <paramref name="angleDegrees"/> around
    /// its center, padding with white.  Output dimensions match input so the
    /// barcode bars remain at the same physical pixel scale.
    /// </summary>
    private static Bitmap RotateBitmap(Bitmap src, double angleDegrees)
    {
        var dst = new Bitmap(src.Width, src.Height, PixelFormat.Format24bppRgb);
        using var g = Graphics.FromImage(dst);

        g.Clear(Color.White);
        g.InterpolationMode  = InterpolationMode.HighQualityBicubic;
        g.SmoothingMode      = SmoothingMode.HighQuality;
        g.PixelOffsetMode    = PixelOffsetMode.HighQuality;

        g.TranslateTransform(src.Width / 2f, src.Height / 2f);
        g.RotateTransform((float)angleDegrees);
        g.TranslateTransform(-src.Width / 2f, -src.Height / 2f);
        g.DrawImage(src, 0, 0);

        return dst;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static Bitmap LoadBitmap(byte[] jpeg)
    {
        using var ms = new MemoryStream(jpeg);
        return new Bitmap(ms);
    }

    private static byte[] ToJpeg(Bitmap bmp, long quality = 90L)
    {
        var encoder = ImageCodecInfo
            .GetImageDecoders()
            .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

        var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);

        using var ms = new MemoryStream();
        bmp.Save(ms, encoder, parameters);
        return ms.ToArray();
    }
}
