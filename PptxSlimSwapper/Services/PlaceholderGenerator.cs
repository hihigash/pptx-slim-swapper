using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Text.Json;

namespace PptxSlimSwapper.Services;

/// <summary>
/// プレースホルダー画像/動画を生成するサービス
/// </summary>
public class PlaceholderGenerator
{
    private const string MetadataKeyword = "pptx_slim_swapper";

    /// <summary>
    /// プレースホルダーに埋め込むメタデータ
    /// </summary>
    public class PlaceholderMetadata
    {
        public string Id { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public string? DataHash { get; set; }
    }
    /// <summary>
    /// プレースホルダー画像を生成する
    /// </summary>
    /// <param name="id">メディアID</param>
    /// <param name="originalFileName">元のファイル名</param>
    /// <param name="contentType">コンテンツタイプ</param>
    /// <param name="dataHash">元データのハッシュ</param>
    /// <returns>画像データのバイト配列</returns>
    public static byte[] GenerateImagePlaceholder(string id, string originalFileName, string contentType, string? dataHash = null)
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
        var pngData = ms.ToArray();

        // メタデータを埋め込む
        return AddMetadataToPng(pngData, id, originalFileName, contentType, dataHash);
    }

    /// <summary>
    /// プレースホルダー動画を生成する(実際には小さい画像を生成)
    /// </summary>
    public static byte[] GenerateVideoPlaceholder(string id, string originalFileName, string contentType, string? dataHash = null)
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
        var pngData = ms.ToArray();

        // メタデータを埋め込む
        return AddMetadataToPng(pngData, id, originalFileName, contentType, dataHash);
    }

    /// <summary>
    /// PNG画像にtEXtチャンクとしてメタデータを追加
    /// </summary>
    private static byte[] AddMetadataToPng(byte[] pngData, string id, string fileName, string contentType, string? dataHash)
    {
        var metadata = new PlaceholderMetadata
        {
            Id = id,
            FileName = fileName,
            ContentType = contentType,
            DataHash = dataHash
        };

        var metadataJson = JsonSerializer.Serialize(metadata);
        var keyword = Encoding.Latin1.GetBytes(MetadataKeyword);
        var text = Encoding.UTF8.GetBytes(metadataJson);

        // tEXtチャンクを構築
        using var ms = new MemoryStream();
        
        // PNG署名とIHDRまでをコピー (最初のIENDチャンクの前まで)
        var iendPos = FindIENDChunk(pngData);
        if (iendPos < 0)
        {
            // IENDが見つからない場合は元のデータをそのまま返す
            return pngData;
        }

        // IENDの前までをコピー
        ms.Write(pngData, 0, iendPos);

        // tEXtチャンクを追加
        WriteTEXtChunk(ms, keyword, text);

        // IENDチャンクを追加
        ms.Write(pngData, iendPos, pngData.Length - iendPos);

        return ms.ToArray();
    }

    /// <summary>
    /// PNG画像からメタデータを読み取る
    /// </summary>
    public static PlaceholderMetadata? ExtractMetadataFromPng(byte[] pngData)
    {
        try
        {
            var chunks = ReadPngChunks(pngData);
            
            foreach (var chunk in chunks)
            {
                if (chunk.Type == "tEXt")
                {
                    // キーワードとテキストを分離 (null区切り)
                    var nullPos = Array.IndexOf(chunk.Data, (byte)0);
                    if (nullPos < 0) continue;

                    var keyword = Encoding.Latin1.GetString(chunk.Data, 0, nullPos);
                    if (keyword == MetadataKeyword)
                    {
                        var text = Encoding.UTF8.GetString(chunk.Data, nullPos + 1, chunk.Data.Length - nullPos - 1);
                        return JsonSerializer.Deserialize<PlaceholderMetadata>(text);
                    }
                }
            }
        }
        catch
        {
            // メタデータ読み取り失敗
        }

        return null;
    }

    private static int FindIENDChunk(byte[] pngData)
    {
        var iendSignature = Encoding.ASCII.GetBytes("IEND");
        for (int i = 8; i < pngData.Length - 12; i++)
        {
            if (pngData[i] == iendSignature[0] &&
                pngData[i + 1] == iendSignature[1] &&
                pngData[i + 2] == iendSignature[2] &&
                pngData[i + 3] == iendSignature[3])
            {
                return i - 4; // チャンク長の位置
            }
        }
        return -1;
    }

    private static void WriteTEXtChunk(Stream stream, byte[] keyword, byte[] text)
    {
        var chunkData = new byte[keyword.Length + 1 + text.Length];
        Array.Copy(keyword, 0, chunkData, 0, keyword.Length);
        chunkData[keyword.Length] = 0; // null separator
        Array.Copy(text, 0, chunkData, keyword.Length + 1, text.Length);

        // チャンク長 (ビッグエンディアン)
        var lengthBytes = BitConverter.GetBytes(chunkData.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(lengthBytes);
        stream.Write(lengthBytes, 0, 4);

        // チャンクタイプ "tEXt"
        var typeBytes = Encoding.ASCII.GetBytes("tEXt");
        stream.Write(typeBytes, 0, 4);

        // チャンクデータ
        stream.Write(chunkData, 0, chunkData.Length);

        // CRC32を計算
        var crc = CalculateCRC32(typeBytes, chunkData);
        var crcBytes = BitConverter.GetBytes(crc);
        if (BitConverter.IsLittleEndian) Array.Reverse(crcBytes);
        stream.Write(crcBytes, 0, 4);
    }

    private static List<(string Type, byte[] Data)> ReadPngChunks(byte[] pngData)
    {
        var chunks = new List<(string, byte[])>();
        var pos = 8; // PNG署名をスキップ

        while (pos < pngData.Length - 12)
        {
            // チャンク長を読み取り
            var lengthBytes = new byte[4];
            Array.Copy(pngData, pos, lengthBytes, 0, 4);
            if (BitConverter.IsLittleEndian) Array.Reverse(lengthBytes);
            var length = BitConverter.ToInt32(lengthBytes, 0);
            pos += 4;

            // チャンクタイプを読み取り
            var type = Encoding.ASCII.GetString(pngData, pos, 4);
            pos += 4;

            // チャンクデータを読み取り
            var data = new byte[length];
            if (length > 0)
            {
                Array.Copy(pngData, pos, data, 0, length);
            }
            pos += length;

            // CRCをスキップ
            pos += 4;

            chunks.Add((type, data));

            if (type == "IEND") break;
        }

        return chunks;
    }

    private static uint CalculateCRC32(byte[] typeBytes, byte[] dataBytes)
    {
        // CRC32テーブルを生成
        var crcTable = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint c = i;
            for (int k = 0; k < 8; k++)
            {
                if ((c & 1) != 0)
                    c = 0xEDB88320 ^ (c >> 1);
                else
                    c >>= 1;
            }
            crcTable[i] = c;
        }

        // CRCを計算
        uint crc = 0xFFFFFFFF;
        
        foreach (var b in typeBytes)
        {
            crc = crcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }
        
        foreach (var b in dataBytes)
        {
            crc = crcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFF;
    }
}
