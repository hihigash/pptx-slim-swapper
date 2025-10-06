using System.Security.Cryptography;

namespace PptxSlimSwapper.Models;

/// <summary>
/// メディアファイル(画像・動画)の情報を保持するクラス
/// </summary>
public class MediaInfo
{
    /// <summary>一意なID(GUID)</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>元のファイル名</summary>
    public string OriginalFileName { get; set; } = string.Empty;

    /// <summary>メディアタイプ(image/video)</summary>
    public string MediaType { get; set; } = string.Empty;

    /// <summary>コンテンツタイプ(image/jpeg, video/mp4など)</summary>
    public string ContentType { get; set; } = string.Empty;

    /// <summary>元のファイルサイズ(バイト)</summary>
    public long OriginalSize { get; set; }

    /// <summary>PPTX内でのパート名</summary>
    public string PartUri { get; set; } = string.Empty;

    /// <summary>保存された外部ファイルパス</summary>
    public string SavedFilePath { get; set; } = string.Empty;

    /// <summary>元のデータのSHA256ハッシュ (復元時の検証用)</summary>
    public string? DataHash { get; set; }

    /// <summary>画像の寸法 (復元時のマッチング精度向上用)</summary>
    public (int Width, int Height)? ImageDimensions { get; set; }

    /// <summary>
    /// データのSHA256ハッシュを計算
    /// </summary>
    public static string ComputeHash(byte[] data)
    {
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(data);
        return Convert.ToBase64String(hash);
    }
}
