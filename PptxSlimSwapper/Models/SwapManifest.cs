namespace PptxSlimSwapper.Models;

/// <summary>
/// 差し替え情報を管理するマニフェストクラス
/// </summary>
public class SwapManifest
{
    /// <summary>作成日時</summary>
    public DateTime CreatedAt { get; set; }

    /// <summary>元のPPTXファイル名</summary>
    public string OriginalFileName { get; set; } = string.Empty;

    /// <summary>差し替えたメディアファイルのリスト</summary>
    public List<MediaInfo> MediaFiles { get; set; } = new();
}
