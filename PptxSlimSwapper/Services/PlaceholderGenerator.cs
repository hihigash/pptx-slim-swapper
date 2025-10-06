using System.Drawing;
using System.Drawing.Imaging;

namespace PptxSlimSwapper.Services;

/// <summary>
/// プレースホルダー画像/動画を生成するサービス
/// </summary>
public class PlaceholderGenerator
{
    /// <summary>
    /// プレースホルダー画像を生成する
    /// </summary>
    /// <param name="id">メディアID</param>
    /// <param name="originalFileName">元のファイル名</param>
    /// <param name="contentType">コンテンツタイプ</param>
    /// <returns>画像データのバイト配列</returns>
    public static byte[] GenerateImagePlaceholder(string id, string originalFileName, string contentType)
    {
        // 小さい画像を生成(100x100ピクセル)
        const int width = 100;
        const int height = 100;

        using var bitmap = new Bitmap(width, height);
        using var graphics = Graphics.FromImage(bitmap);

        // 背景を薄いグレーで塗りつぶし
        graphics.Clear(Color.LightGray);

        // 境界線を描画
        using var pen = new Pen(Color.DarkGray, 2);
        graphics.DrawRectangle(pen, 1, 1, width - 2, height - 2);

        // テキストを描画
        using var font = new Font("Arial", 8);
        using var brush = new SolidBrush(Color.Black);
        var text = $"[PLACEHOLDER]\n{Path.GetFileName(originalFileName)}\nID: {id.Substring(0, 8)}...";
        var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };
        graphics.DrawString(text, font, brush, new RectangleF(0, 0, width, height), format);

        // PNGとして保存
        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>
    /// プレースホルダー動画を生成する(実際には小さい画像を生成)
    /// </summary>
    public static byte[] GenerateVideoPlaceholder(string id, string originalFileName, string contentType)
    {
        // 動画も同様に小さい画像として生成
        const int width = 100;
        const int height = 100;

        using var bitmap = new Bitmap(width, height);
        using var graphics = Graphics.FromImage(bitmap);

        // 背景を薄い青で塗りつぶし
        graphics.Clear(Color.LightBlue);

        // 境界線を描画
        using var pen = new Pen(Color.DarkBlue, 2);
        graphics.DrawRectangle(pen, 1, 1, width - 2, height - 2);

        // 再生ボタン風の三角形を描画
        var trianglePoints = new[]
        {
            new Point(40, 30),
            new Point(40, 70),
            new Point(70, 50)
        };
        using var blueBrush = new SolidBrush(Color.DarkBlue);
        graphics.FillPolygon(blueBrush, trianglePoints);

        // テキストを描画
        using var font = new Font("Arial", 6);
        using var brush = new SolidBrush(Color.Black);
        var text = $"[VIDEO]\n{Path.GetFileName(originalFileName)}\n{id.Substring(0, 8)}";
        var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Near
        };
        graphics.DrawString(text, font, brush, new RectangleF(0, 75, width, 25), format);

        // PNGとして保存
        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }
}
