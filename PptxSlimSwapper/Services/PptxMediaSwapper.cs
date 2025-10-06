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

        // データのハッシュを計算
        var dataHash = Models.MediaInfo.ComputeHash(originalData);

        // 画像の寸法を取得 (可能な場合)
        (int Width, int Height)? dimensions = null;
        try
        {
            using var ms = new MemoryStream(originalData);
            using var image = System.Drawing.Image.FromStream(ms);
            dimensions = (image.Width, image.Height);
        }
        catch
        {
            // 寸法取得失敗時は無視
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
            SavedFilePath = Path.Combine(MediaFolderName, $"{Guid.NewGuid()}{GetExtensionFromContentType(imagePart.ContentType)}"),
            DataHash = dataHash,
            ImageDimensions = dimensions
        };

        // 元の画像を保存
        var savedFilePath = Path.Combine(Path.GetDirectoryName(mediaDirectory) ?? "", mediaInfo.SavedFilePath);
        await File.WriteAllBytesAsync(savedFilePath, originalData);

        // プレースホルダー画像を生成して差し替え
        var placeholderData = PlaceholderGenerator.GenerateImagePlaceholder(
            mediaInfo.Id,
            mediaInfo.OriginalFileName,
            mediaInfo.ContentType,
            mediaInfo.DataHash
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

        // データのハッシュを計算
        var dataHash = Models.MediaInfo.ComputeHash(originalData);

        // メディア情報を作成
        var mediaInfo = new MediaInfo
        {
            Id = Guid.NewGuid().ToString(),
            OriginalFileName = GetFileNameFromUri(videoPart.Uri?.ToString() ?? "unknown.mp4"),
            MediaType = "video",
            ContentType = videoPart.ContentType,
            OriginalSize = originalData.Length,
            PartUri = videoPart.Uri?.ToString() ?? "",
            SavedFilePath = Path.Combine(MediaFolderName, $"{Guid.NewGuid()}{GetExtensionFromContentType(videoPart.ContentType)}"),
            DataHash = dataHash
        };

        // 元の動画を保存
        var savedFilePath = Path.Combine(Path.GetDirectoryName(mediaDirectory) ?? "", mediaInfo.SavedFilePath);
        await File.WriteAllBytesAsync(savedFilePath, originalData);

        // プレースホルダー画像を生成して差し替え(動画の場合も画像で代替)
        var placeholderData = PlaceholderGenerator.GenerateVideoPlaceholder(
            mediaInfo.Id,
            mediaInfo.OriginalFileName,
            mediaInfo.ContentType,
            mediaInfo.DataHash
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

        // 柔軟なマッチングで画像パーツを検索
        var targetPart = await FindImagePartByFlexibleMatchingAsync(allImageParts, mediaInfo);

        if (targetPart == null)
        {
            Console.WriteLine($"警告: 画像パーツが見つかりません: {mediaInfo.PartUri} (ID: {mediaInfo.Id})");
            return;
        }

        // 保存されていた元の画像データを読み込む
        var originalData = await File.ReadAllBytesAsync(savedFilePath);

        // 復元
        using var stream = new MemoryStream(originalData);
        targetPart.FeedData(stream);
        
        Console.WriteLine($"復元成功: {mediaInfo.OriginalFileName} (ID: {mediaInfo.Id})");
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

        // 柔軟なマッチングで動画パーツを検索
        var targetPart = await FindVideoPartByFlexibleMatchingAsync(allVideoParts, mediaInfo);

        if (targetPart == null)
        {
            Console.WriteLine($"警告: 動画パーツが見つかりません: {mediaInfo.PartUri} (ID: {mediaInfo.Id})");
            return;
        }

        // 保存されていた元の動画データを読み込む
        var originalData = await File.ReadAllBytesAsync(savedFilePath);

        // 復元
        using var stream = new MemoryStream(originalData);
        targetPart.FeedData(stream);
        
        Console.WriteLine($"復元成功: {mediaInfo.OriginalFileName} (ID: {mediaInfo.Id})");
    }

    private static string GetFileNameFromUri(string uri)
    {
        return Path.GetFileName(uri.Replace('/', Path.DirectorySeparatorChar));
    }

    /// <summary>
    /// 柔軟なマッチング戦略で画像パーツを検索
    /// </summary>
    private static async Task<ImagePart?> FindImagePartByFlexibleMatchingAsync(
        List<ImagePart> allImageParts,
        MediaInfo mediaInfo)
    {
        // 重複を除外
        var uniqueParts = allImageParts
            .GroupBy(ip => ip.Uri?.ToString())
            .Select(g => g.First())
            .ToList();

        // 戦略1: URI完全一致 (最優先)
        var exactMatch = uniqueParts.FirstOrDefault(ip => ip.Uri?.ToString() == mediaInfo.PartUri);
        if (exactMatch != null)
        {
            Console.WriteLine($"  → URI完全一致でマッチング: {mediaInfo.PartUri}");
            return exactMatch;
        }

        // 戦略2: プレースホルダーのメタデータからIDを照合
        Console.WriteLine($"  → URI一致なし。メタデータでマッチング試行中...");
        foreach (var part in uniqueParts.Where(p => p.ContentType.StartsWith("image/")))
        {
            try
            {
                using var stream = part.GetStream();
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                var imageData = ms.ToArray();

                var metadata = PlaceholderGenerator.ExtractMetadataFromPng(imageData);
                if (metadata != null && metadata.Id == mediaInfo.Id)
                {
                    Console.WriteLine($"  → メタデータID一致でマッチング: {mediaInfo.Id}");
                    return part;
                }
            }
            catch
            {
                // メタデータ読み取り失敗は無視
            }
        }

        // 戦略3: ContentTypeとファイル名の組み合わせ
        Console.WriteLine($"  → ファイル名とContentTypeでマッチング試行中...");
        var originalName = Path.GetFileNameWithoutExtension(mediaInfo.OriginalFileName);
        var similarMatch = uniqueParts
            .Where(ip => ip.ContentType == mediaInfo.ContentType)
            .FirstOrDefault(ip =>
            {
                var fileName = GetFileNameFromUri(ip.Uri?.ToString() ?? "");
                return Path.GetFileNameWithoutExtension(fileName)
                    .Contains(originalName, StringComparison.OrdinalIgnoreCase);
            });

        if (similarMatch != null)
        {
            Console.WriteLine($"  → ファイル名類似性でマッチング: {GetFileNameFromUri(similarMatch.Uri?.ToString() ?? "")}");
        }

        return similarMatch;
    }

    /// <summary>
    /// 柔軟なマッチング戦略で動画パーツを検索
    /// </summary>
    private static async Task<DataPart?> FindVideoPartByFlexibleMatchingAsync(
        List<DataPart> allVideoParts,
        MediaInfo mediaInfo)
    {
        // 重複を除外
        var uniqueParts = allVideoParts
            .GroupBy(vp => vp.Uri?.ToString())
            .Select(g => g.First())
            .ToList();

        // 戦略1: URI完全一致 (最優先)
        var exactMatch = uniqueParts.FirstOrDefault(vp => vp.Uri?.ToString() == mediaInfo.PartUri);
        if (exactMatch != null)
        {
            Console.WriteLine($"  → URI完全一致でマッチング: {mediaInfo.PartUri}");
            return exactMatch;
        }

        // 戦略2: プレースホルダーのメタデータからIDを照合 (動画もPNG画像で代替しているため)
        Console.WriteLine($"  → URI一致なし。メタデータでマッチング試行中...");
        foreach (var part in uniqueParts)
        {
            try
            {
                using var stream = part.GetStream();
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                var data = ms.ToArray();

                // PNG形式かチェック
                if (data.Length > 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                {
                    var metadata = PlaceholderGenerator.ExtractMetadataFromPng(data);
                    if (metadata != null && metadata.Id == mediaInfo.Id)
                    {
                        Console.WriteLine($"  → メタデータID一致でマッチング: {mediaInfo.Id}");
                        return part;
                    }
                }
            }
            catch
            {
                // メタデータ読み取り失敗は無視
            }
        }

        // 戦略3: ContentTypeとファイル名の組み合わせ
        Console.WriteLine($"  → ファイル名とContentTypeでマッチング試行中...");
        var originalName = Path.GetFileNameWithoutExtension(mediaInfo.OriginalFileName);
        var similarMatch = uniqueParts
            .Where(vp => vp.ContentType == mediaInfo.ContentType)
            .FirstOrDefault(vp =>
            {
                var fileName = GetFileNameFromUri(vp.Uri?.ToString() ?? "");
                return Path.GetFileNameWithoutExtension(fileName)
                    .Contains(originalName, StringComparison.OrdinalIgnoreCase);
            });

        if (similarMatch != null)
        {
            Console.WriteLine($"  → ファイル名類似性でマッチング: {GetFileNameFromUri(similarMatch.Uri?.ToString() ?? "")}");
        }

        return similarMatch;
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
