using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using PptxSlimSwapper.Models;

namespace PptxSlimSwapper.Services;

/// <summary>
/// PPTXファイルからメディアを抽出して差し替えるサービス
/// </summary>
public class PptxMediaSwapper
{
    private const string ManifestFileName = "swap-manifest.json";
    private const string MediaFolderName = "media";

    /// <summary>
    /// PPTXファイル内のメディアを外部に保存し、プレースホルダーに差し替える
    /// </summary>
    /// <param name="pptxFilePath">PPTXファイルのパス</param>
    /// <param name="outputDirectory">出力ディレクトリ</param>
    public static async Task<(string outputPptxPath, string manifestPath, int mediaCount)> SwapOutMediaAsync(
        string pptxFilePath, 
        string outputDirectory)
    {
        if (!File.Exists(pptxFilePath))
        {
            throw new FileNotFoundException("PPTXファイルが見つかりません。", pptxFilePath);
        }

        // 出力ディレクトリの準備
        Directory.CreateDirectory(outputDirectory);
        var mediaDirectory = Path.Combine(outputDirectory, MediaFolderName);
        Directory.CreateDirectory(mediaDirectory);

        var manifest = new SwapManifest
        {
            CreatedAt = DateTime.Now,
            OriginalFileName = Path.GetFileName(pptxFilePath)
        };

        // 出力PPTXファイルのパス
        var outputPptxPath = Path.Combine(
            outputDirectory,
            $"{Path.GetFileNameWithoutExtension(pptxFilePath)}_slim{Path.GetExtension(pptxFilePath)}"
        );

        // PPTXファイルをコピー
        File.Copy(pptxFilePath, outputPptxPath, overwrite: true);

        int mediaCount = 0;

        // PPTXファイルを開いて処理
        using (var presentationDocument = PresentationDocument.Open(outputPptxPath, isEditable: true))
        {
            var allImageParts = new List<ImagePart>();
            var allVideoParts = new List<DataPart>();

            if (presentationDocument.PresentationPart != null)
            {
                // 通常のスライドから画像を取得
                if (presentationDocument.PresentationPart.SlideParts != null)
                {
                    allImageParts.AddRange(
                        presentationDocument.PresentationPart.SlideParts
                            .SelectMany(sp => sp.ImageParts));
                    
                    allVideoParts.AddRange(
                        presentationDocument.PresentationPart.SlideParts
                            .SelectMany(sp => sp.DataPartReferenceRelationships
                                .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                                .Select(r => r.DataPart)));
                }

                // スライドマスタから画像を取得
                if (presentationDocument.PresentationPart.SlideMasterParts != null)
                {
                    allImageParts.AddRange(
                        presentationDocument.PresentationPart.SlideMasterParts
                            .SelectMany(smp => smp.ImageParts));
                    
                    allVideoParts.AddRange(
                        presentationDocument.PresentationPart.SlideMasterParts
                            .SelectMany(smp => smp.DataPartReferenceRelationships
                                .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                                .Select(r => r.DataPart)));

                    // スライドマスタ配下のレイアウトスライドから画像を取得
                    foreach (var slideMasterPart in presentationDocument.PresentationPart.SlideMasterParts)
                    {
                        if (slideMasterPart.SlideLayoutParts != null)
                        {
                            allImageParts.AddRange(
                                slideMasterPart.SlideLayoutParts
                                    .SelectMany(slp => slp.ImageParts));
                            
                            allVideoParts.AddRange(
                                slideMasterPart.SlideLayoutParts
                                    .SelectMany(slp => slp.DataPartReferenceRelationships
                                        .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                                        .Select(r => r.DataPart)));
                        }
                    }
                }

                // ノートマスタから画像を取得
                if (presentationDocument.PresentationPart.NotesMasterPart != null)
                {
                    allImageParts.AddRange(presentationDocument.PresentationPart.NotesMasterPart.ImageParts);
                    
                    allVideoParts.AddRange(
                        presentationDocument.PresentationPart.NotesMasterPart.DataPartReferenceRelationships
                            .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                            .Select(r => r.DataPart));
                }

                // ハンドアウトマスタから画像を取得
                if (presentationDocument.PresentationPart.HandoutMasterPart != null)
                {
                    allImageParts.AddRange(presentationDocument.PresentationPart.HandoutMasterPart.ImageParts);
                    
                    allVideoParts.AddRange(
                        presentationDocument.PresentationPart.HandoutMasterPart.DataPartReferenceRelationships
                            .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                            .Select(r => r.DataPart));
                }
            }

            // 画像パーツの処理 - URIベースで重複除外
            var imageParts = allImageParts
                .GroupBy(ip => ip.Uri?.ToString())
                .Select(g => g.First())
                .ToList();

            Console.WriteLine($"処理する画像パーツ数: {imageParts.Count}");

            foreach (var imagePart in imageParts)
            {
                await ProcessImagePartAsync(imagePart, mediaDirectory, manifest);
                mediaCount++;
            }

            // 動画パーツの処理 - URIベースで重複除外
            var videoParts = allVideoParts
                .GroupBy(vp => vp.Uri?.ToString())
                .Select(g => g.First())
                .ToList();

            Console.WriteLine($"処理する動画パーツ数: {videoParts.Count}");

            foreach (var videoPart in videoParts)
            {
                await ProcessVideoPartAsync(videoPart, mediaDirectory, manifest);
                mediaCount++;
            }
        }

        // マニフェストファイルを保存
        var manifestPath = Path.Combine(outputDirectory, ManifestFileName);
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        var manifestJson = JsonSerializer.Serialize(manifest, jsonOptions);
        await File.WriteAllTextAsync(manifestPath, manifestJson);

        return (outputPptxPath, manifestPath, mediaCount);
    }

    /// <summary>
    /// プレースホルダーに差し替えたメディアを元に戻す
    /// </summary>
    /// <param name="slimPptxFilePath">差し替え済みPPTXファイルのパス</param>
    /// <param name="manifestDirectory">マニフェストファイルがあるディレクトリ</param>
    public static async Task<(string outputPptxPath, int restoredCount)> SwapInMediaAsync(
        string slimPptxFilePath,
        string manifestDirectory)
    {
        var manifestPath = Path.Combine(manifestDirectory, ManifestFileName);
        if (!File.Exists(manifestPath))
        {
            throw new FileNotFoundException("マニフェストファイルが見つかりません。", manifestPath);
        }

        // マニフェストを読み込む
        var manifestJson = await File.ReadAllTextAsync(manifestPath);
        var manifest = JsonSerializer.Deserialize<SwapManifest>(manifestJson)
            ?? throw new InvalidOperationException("マニフェストの読み込みに失敗しました。");

        // 出力PPTXファイルのパス
        var outputPptxPath = Path.Combine(
            Path.GetDirectoryName(slimPptxFilePath) ?? "",
            $"{Path.GetFileNameWithoutExtension(slimPptxFilePath)}_restored{Path.GetExtension(slimPptxFilePath)}"
        );

        // PPTXファイルをコピー
        File.Copy(slimPptxFilePath, outputPptxPath, overwrite: true);

        int restoredCount = 0;

        // PPTXファイルを開いて処理
        using (var presentationDocument = PresentationDocument.Open(outputPptxPath, isEditable: true))
        {
            foreach (var mediaInfo in manifest.MediaFiles)
            {
                var savedFilePath = Path.Combine(manifestDirectory, mediaInfo.SavedFilePath);
                if (!File.Exists(savedFilePath))
                {
                    Console.WriteLine($"警告: メディアファイルが見つかりません: {savedFilePath}");
                    continue;
                }

                if (mediaInfo.MediaType == "image")
                {
                    await RestoreImagePartAsync(presentationDocument, mediaInfo, savedFilePath);
                    restoredCount++;
                }
                else if (mediaInfo.MediaType == "video")
                {
                    await RestoreVideoPartAsync(presentationDocument, mediaInfo, savedFilePath);
                    restoredCount++;
                }
            }
        }

        return (outputPptxPath, restoredCount);
    }

    private static async Task ProcessImagePartAsync(ImagePart imagePart, string mediaDirectory, SwapManifest manifest)
    {
        // 元の画像データを読み込む
        byte[] originalData;
        using (var stream = imagePart.GetStream())
        using (var memoryStream = new MemoryStream())
        {
            await stream.CopyToAsync(memoryStream);
            originalData = memoryStream.ToArray();
        }

        // メディア情報を作成
        var mediaInfo = new MediaInfo
        {
            Id = Guid.NewGuid().ToString(),
            OriginalFileName = GetFileNameFromUri(imagePart.Uri?.ToString() ?? "unknown.png"),
            MediaType = "image",
            ContentType = imagePart.ContentType,
            OriginalSize = originalData.Length,
            PartUri = imagePart.Uri?.ToString() ?? "",
            SavedFilePath = Path.Combine(MediaFolderName, $"{Guid.NewGuid()}{GetExtensionFromContentType(imagePart.ContentType)}")
        };

        // 元の画像を保存
        var savedFilePath = Path.Combine(Path.GetDirectoryName(mediaDirectory) ?? "", mediaInfo.SavedFilePath);
        await File.WriteAllBytesAsync(savedFilePath, originalData);

        // プレースホルダー画像を生成して差し替え
        var placeholderData = PlaceholderGenerator.GenerateImagePlaceholder(
            mediaInfo.Id,
            mediaInfo.OriginalFileName,
            mediaInfo.ContentType
        );

        using var placeholderStream = new MemoryStream(placeholderData);
        imagePart.FeedData(placeholderStream);

        manifest.MediaFiles.Add(mediaInfo);
    }

    private static async Task ProcessVideoPartAsync(DataPart videoPart, string mediaDirectory, SwapManifest manifest)
    {
        // 元の動画データを読み込む
        byte[] originalData;
        using (var stream = videoPart.GetStream())
        using (var memoryStream = new MemoryStream())
        {
            await stream.CopyToAsync(memoryStream);
            originalData = memoryStream.ToArray();
        }

        // メディア情報を作成
        var mediaInfo = new MediaInfo
        {
            Id = Guid.NewGuid().ToString(),
            OriginalFileName = GetFileNameFromUri(videoPart.Uri?.ToString() ?? "unknown.mp4"),
            MediaType = "video",
            ContentType = videoPart.ContentType,
            OriginalSize = originalData.Length,
            PartUri = videoPart.Uri?.ToString() ?? "",
            SavedFilePath = Path.Combine(MediaFolderName, $"{Guid.NewGuid()}{GetExtensionFromContentType(videoPart.ContentType)}")
        };

        // 元の動画を保存
        var savedFilePath = Path.Combine(Path.GetDirectoryName(mediaDirectory) ?? "", mediaInfo.SavedFilePath);
        await File.WriteAllBytesAsync(savedFilePath, originalData);

        // プレースホルダー画像を生成して差し替え(動画の場合も画像で代替)
        var placeholderData = PlaceholderGenerator.GenerateVideoPlaceholder(
            mediaInfo.Id,
            mediaInfo.OriginalFileName,
            mediaInfo.ContentType
        );

        using var placeholderStream = new MemoryStream(placeholderData);
        videoPart.FeedData(placeholderStream);

        manifest.MediaFiles.Add(mediaInfo);
    }

    private static async Task RestoreImagePartAsync(
        PresentationDocument presentationDocument,
        MediaInfo mediaInfo,
        string savedFilePath)
    {
        var allImageParts = new List<ImagePart>();

        if (presentationDocument.PresentationPart != null)
        {
            // 通常のスライドから画像パーツを取得
            if (presentationDocument.PresentationPart.SlideParts != null)
            {
                allImageParts.AddRange(
                    presentationDocument.PresentationPart.SlideParts
                        .SelectMany(sp => sp.ImageParts));
            }

            // スライドマスタから画像パーツを取得
            if (presentationDocument.PresentationPart.SlideMasterParts != null)
            {
                allImageParts.AddRange(
                    presentationDocument.PresentationPart.SlideMasterParts
                        .SelectMany(smp => smp.ImageParts));

                // スライドマスタ配下のレイアウトスライドから画像を取得
                foreach (var slideMasterPart in presentationDocument.PresentationPart.SlideMasterParts)
                {
                    if (slideMasterPart.SlideLayoutParts != null)
                    {
                        allImageParts.AddRange(
                            slideMasterPart.SlideLayoutParts
                                .SelectMany(slp => slp.ImageParts));
                    }
                }
            }

            // ノートマスタから画像を取得
            if (presentationDocument.PresentationPart.NotesMasterPart != null)
            {
                allImageParts.AddRange(presentationDocument.PresentationPart.NotesMasterPart.ImageParts);
            }

            // ハンドアウトマスタから画像を取得
            if (presentationDocument.PresentationPart.HandoutMasterPart != null)
            {
                allImageParts.AddRange(presentationDocument.PresentationPart.HandoutMasterPart.ImageParts);
            }
        }

        // 画像パーツを検索 - URI基準で重複除外してから検索
        var imageParts = allImageParts
            .GroupBy(ip => ip.Uri?.ToString())
            .Where(g => g.Key == mediaInfo.PartUri)
            .SelectMany(g => g) // グループ内の全てのパーツを取得
            .ToList();

        if (imageParts.Count == 0)
        {
            Console.WriteLine($"警告: 画像パーツが見つかりません: {mediaInfo.PartUri}");
            return;
        }

        // 保存されていた元の画像データを読み込む
        var originalData = await File.ReadAllBytesAsync(savedFilePath);

        // 該当するパーツの最初の1つだけに復元（同じURIは同じ実体を指す）
        var firstPart = imageParts.First();
        using var stream = new MemoryStream(originalData);
        firstPart.FeedData(stream);
    }

    private static async Task RestoreVideoPartAsync(
        PresentationDocument presentationDocument,
        MediaInfo mediaInfo,
        string savedFilePath)
    {
        var allVideoParts = new List<DataPart>();

        if (presentationDocument.PresentationPart != null)
        {
            // 通常のスライドから動画パーツを取得
            if (presentationDocument.PresentationPart.SlideParts != null)
            {
                allVideoParts.AddRange(
                    presentationDocument.PresentationPart.SlideParts
                        .SelectMany(sp => sp.DataPartReferenceRelationships
                            .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                            .Select(r => r.DataPart)));
            }

            // スライドマスタから動画パーツを取得
            if (presentationDocument.PresentationPart.SlideMasterParts != null)
            {
                allVideoParts.AddRange(
                    presentationDocument.PresentationPart.SlideMasterParts
                        .SelectMany(smp => smp.DataPartReferenceRelationships
                            .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                            .Select(r => r.DataPart)));

                // スライドマスタ配下のレイアウトスライドから動画を取得
                foreach (var slideMasterPart in presentationDocument.PresentationPart.SlideMasterParts)
                {
                    if (slideMasterPart.SlideLayoutParts != null)
                    {
                        allVideoParts.AddRange(
                            slideMasterPart.SlideLayoutParts
                                .SelectMany(slp => slp.DataPartReferenceRelationships
                                    .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                                    .Select(r => r.DataPart)));
                    }
                }
            }

            // ノートマスタから動画を取得
            if (presentationDocument.PresentationPart.NotesMasterPart != null)
            {
                allVideoParts.AddRange(
                    presentationDocument.PresentationPart.NotesMasterPart.DataPartReferenceRelationships
                        .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                        .Select(r => r.DataPart));
            }

            // ハンドアウトマスタから動画を取得
            if (presentationDocument.PresentationPart.HandoutMasterPart != null)
            {
                allVideoParts.AddRange(
                    presentationDocument.PresentationPart.HandoutMasterPart.DataPartReferenceRelationships
                        .Where(r => r.DataPart.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
                        .Select(r => r.DataPart));
            }
        }

        // 動画パーツを検索 - URI基準で重複除外してから検索
        var videoParts = allVideoParts
            .GroupBy(vp => vp.Uri?.ToString())
            .Where(g => g.Key == mediaInfo.PartUri)
            .SelectMany(g => g) // グループ内の全てのパーツを取得
            .ToList();

        if (videoParts.Count == 0)
        {
            Console.WriteLine($"警告: 動画パーツが見つかりません: {mediaInfo.PartUri}");
            return;
        }

        // 保存されていた元の動画データを読み込む
        var originalData = await File.ReadAllBytesAsync(savedFilePath);

        // 該当するパーツの最初の1つだけに復元（同じURIは同じ実体を指す）
        var firstPart = videoParts.First();
        using var stream = new MemoryStream(originalData);
        firstPart.FeedData(stream);
    }

    private static string GetFileNameFromUri(string uri)
    {
        return Path.GetFileName(uri.Replace('/', Path.DirectorySeparatorChar));
    }

    private static string GetExtensionFromContentType(string contentType)
    {
        return contentType.ToLowerInvariant() switch
        {
            "image/jpeg" => ".jpg",
            "image/png" => ".png",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "video/mp4" => ".mp4",
            "video/avi" => ".avi",
            "video/wmv" => ".wmv",
            "video/mov" => ".mov",
            _ => ".dat"
        };
    }
}
