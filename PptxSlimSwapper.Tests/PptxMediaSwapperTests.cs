using PptxSlimSwapper.Services;

namespace PptxSlimSwapper.Tests;

public class PptxMediaSwapperTests
{
    private string _testDataDirectory = string.Empty;
    private string _tempDirectory = string.Empty;

    [SetUp]
    public void Setup()
    {
        // テストデータのディレクトリ (data/)
        var currentDir = Directory.GetCurrentDirectory();
        var projectRoot = Path.GetFullPath(Path.Combine(currentDir, "..", "..", "..", ".."));
        _testDataDirectory = Path.Combine(projectRoot, "data");
        
        // 一時作業ディレクトリ
        _tempDirectory = Path.Combine(Path.GetTempPath(), $"PptxSlimSwapper_Test_{Guid.NewGuid()}");
        Directory.CreateDirectory(_tempDirectory);
    }

    [TearDown]
    public void TearDown()
    {
        // 一時ディレクトリをクリーンアップ
        if (Directory.Exists(_tempDirectory))
        {
            try
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
            catch
            {
                // クリーンアップに失敗しても無視
            }
        }
    }

    /// <summary>
    /// テストデータディレクトリ内のすべてのPPTXファイルを取得
    /// </summary>
    private List<string> GetTestPptxFiles()
    {
        if (!Directory.Exists(_testDataDirectory))
        {
            return new List<string>();
        }

        return Directory.GetFiles(_testDataDirectory, "*.pptx", SearchOption.TopDirectoryOnly)
            .ToList();
    }

    [Test]
    public async Task OutCommand_ShouldReduceFileSize()
    {
        // Arrange
        var testFiles = GetTestPptxFiles();
        
        if (testFiles.Count == 0)
        {
            Assert.Ignore($"テストファイルが見つかりません。{_testDataDirectory} ディレクトリに .pptx ファイルを配置してください。");
            return;
        }

        foreach (var testPptxPath in testFiles)
        {
            Console.WriteLine($"\n=== テスト対象: {Path.GetFileName(testPptxPath)} ===");
            
            var outputDirectory = Path.Combine(_tempDirectory, Path.GetFileNameWithoutExtension(testPptxPath) + "_output");

            // Act
            var (outputPptxPath, manifestPath, mediaCount) = await PptxMediaSwapper.SwapOutMediaAsync(
                testPptxPath,
                outputDirectory
            );

            // Assert
            Assert.That(File.Exists(outputPptxPath), Is.True, $"出力PPTXファイルが作成されていません: {testPptxPath}");
            Assert.That(File.Exists(manifestPath), Is.True, $"マニフェストファイルが作成されていません: {testPptxPath}");

            var originalSize = new FileInfo(testPptxPath).Length;
            var slimSize = new FileInfo(outputPptxPath).Length;
            
            Assert.That(slimSize, Is.LessThanOrEqualTo(originalSize), $"ファイルサイズが増加しています: {testPptxPath}");
            
            var reduction = originalSize - slimSize;
            var reductionPercent = originalSize > 0 ? (double)reduction / originalSize * 100 : 0;

            Console.WriteLine($"元のサイズ: {FormatFileSize(originalSize)}");
            Console.WriteLine($"差し替え後のサイズ: {FormatFileSize(slimSize)}");
            Console.WriteLine($"削減量: {FormatFileSize(reduction)} ({reductionPercent:F2}%)");
            Console.WriteLine($"差し替えたメディア数: {mediaCount}");

            if (mediaCount == 0)
            {
                Console.WriteLine("⚠ このファイルにはメディアが含まれていません");
            }
        }

        Assert.Pass($"✓ {testFiles.Count} 個のファイルをテストしました");
    }

    [Test]
    public async Task InCommand_ShouldRestoreOriginalFileSize()
    {
        // Arrange
        var testFiles = GetTestPptxFiles();
        
        if (testFiles.Count == 0)
        {
            Assert.Ignore($"テストファイルが見つかりません。{_testDataDirectory} ディレクトリに .pptx ファイルを配置してください。");
            return;
        }

        foreach (var testPptxPath in testFiles)
        {
            Console.WriteLine($"\n=== テスト対象: {Path.GetFileName(testPptxPath)} ===");
            
            var outputDirectory = Path.Combine(_tempDirectory, Path.GetFileNameWithoutExtension(testPptxPath) + "_output");

            // まず差し替えを実行
            var (slimPptxPath, manifestPath, mediaCount) = await PptxMediaSwapper.SwapOutMediaAsync(
                testPptxPath,
                outputDirectory
            );

            Assert.That(File.Exists(slimPptxPath), Is.True, $"差し替え済みPPTXファイルが作成されていません: {testPptxPath}");

            var originalSize = new FileInfo(testPptxPath).Length;
            var slimSize = new FileInfo(slimPptxPath).Length;

            if (mediaCount == 0)
            {
                Console.WriteLine("⚠ このファイルにはメディアが含まれていないため、復元テストをスキップします");
                continue;
            }

            // Act - 復元を実行
            var (restoredPptxPath, restoredCount) = await PptxMediaSwapper.SwapInMediaAsync(
                slimPptxPath,
                outputDirectory
            );

            // Assert
            Assert.That(File.Exists(restoredPptxPath), Is.True, $"復元されたPPTXファイルが作成されていません: {testPptxPath}");
            Assert.That(restoredCount, Is.EqualTo(mediaCount), $"復元されたメディア数が一致しません: {testPptxPath}");

            var restoredSize = new FileInfo(restoredPptxPath).Length;

            // ファイルサイズが元のサイズと非常に近いことを確認（OpenXMLの圧縮やメタデータの違いで若干異なる可能性がある）
            var sizeDifference = Math.Abs(restoredSize - originalSize);
            var differencePercent = originalSize > 0 ? (double)sizeDifference / originalSize * 100 : 0;
            
            // 差異が1%以内であることを確認
            Assert.That(differencePercent, Is.LessThan(1.0), 
                $"復元後のファイルサイズの差異が大きすぎます ({Path.GetFileName(testPptxPath)}): 元: {originalSize:N0}, 復元後: {restoredSize:N0}, 差異: {sizeDifference:N0} ({differencePercent:F2}%)");

            Console.WriteLine($"元のサイズ: {FormatFileSize(originalSize)}");
            Console.WriteLine($"差し替え後のサイズ: {FormatFileSize(slimSize)}");
            Console.WriteLine($"復元後のサイズ: {FormatFileSize(restoredSize)}");
            Console.WriteLine($"サイズ差異: {FormatFileSize(sizeDifference)} ({differencePercent:F4}%)");
            Console.WriteLine($"差し替え/復元したメディア数: {mediaCount} / {restoredCount}");
            Console.WriteLine($"✓ ファイルサイズがほぼ一致しました（差異: {differencePercent:F4}%）");
        }

        Assert.Pass($"✓ {testFiles.Count} 個のファイルをテストしました");
    }

    [Test]
    public async Task OutAndIn_RoundTrip_ShouldMaintainFileSize()
    {
        // Arrange
        var testFiles = GetTestPptxFiles();
        
        if (testFiles.Count == 0)
        {
            Assert.Ignore($"テストファイルが見つかりません。{_testDataDirectory} ディレクトリに .pptx ファイルを配置してください。");
            return;
        }

        var successCount = 0;
        var skippedCount = 0;

        foreach (var testPptxPath in testFiles)
        {
            Console.WriteLine($"\n=== Round Trip Test: {Path.GetFileName(testPptxPath)} ===");
            
            var outputDirectory = Path.Combine(_tempDirectory, Path.GetFileNameWithoutExtension(testPptxPath) + "_output");
            var originalSize = new FileInfo(testPptxPath).Length;

            // Act - Out
            var (slimPptxPath, _, mediaCount) = await PptxMediaSwapper.SwapOutMediaAsync(
                testPptxPath,
                outputDirectory
            );

            var slimSize = new FileInfo(slimPptxPath).Length;
            var reduction = originalSize - slimSize;
            var reductionPercent = originalSize > 0 ? (double)reduction / originalSize * 100 : 0;

            if (mediaCount == 0)
            {
                Console.WriteLine("⚠ このファイルにはメディアが含まれていないため、Round Tripテストをスキップします");
                skippedCount++;
                continue;
            }

            // Act - In
            var (restoredPptxPath, restoredCount) = await PptxMediaSwapper.SwapInMediaAsync(
                slimPptxPath,
                outputDirectory
            );

            var restoredSize = new FileInfo(restoredPptxPath).Length;

            // ファイルサイズの差異を計算
            var sizeDifference = Math.Abs(restoredSize - originalSize);
            var differencePercent = originalSize > 0 ? (double)sizeDifference / originalSize * 100 : 0;

            // Assert
            Assert.Multiple(() =>
            {
                Assert.That(mediaCount, Is.GreaterThan(0), $"メディアファイルが見つかりませんでした: {Path.GetFileName(testPptxPath)}");
                Assert.That(slimSize, Is.LessThanOrEqualTo(originalSize), $"ファイルサイズが増加しています: {Path.GetFileName(testPptxPath)}");
                Assert.That(restoredCount, Is.EqualTo(mediaCount), $"復元されたメディア数が一致しません: {Path.GetFileName(testPptxPath)}");
                
                // 差異が1%以内であることを確認
                Assert.That(differencePercent, Is.LessThan(1.0), 
                    $"復元後のファイルサイズの差異が大きすぎます ({Path.GetFileName(testPptxPath)}): 元: {originalSize:N0}, 復元後: {restoredSize:N0}, 差異: {sizeDifference:N0} ({differencePercent:F2}%)");
            });

            Console.WriteLine($"元のサイズ: {FormatFileSize(originalSize)}");
            Console.WriteLine($"差し替え後のサイズ: {FormatFileSize(slimSize)}");
            Console.WriteLine($"削減量: {FormatFileSize(reduction)} ({reductionPercent:F2}%)");
            Console.WriteLine($"復元後のサイズ: {FormatFileSize(restoredSize)}");
            Console.WriteLine($"サイズ差異: {FormatFileSize(sizeDifference)} ({differencePercent:F4}%)");
            Console.WriteLine($"処理したメディア数: {mediaCount}");
            Console.WriteLine($"✓ ファイルサイズがほぼ一致しました（差異: {differencePercent:F4}%）");

            successCount++;
        }

        Console.WriteLine($"\n=== テスト結果サマリー ===");
        Console.WriteLine($"成功: {successCount} 個");
        Console.WriteLine($"スキップ: {skippedCount} 個");
        Console.WriteLine($"合計: {testFiles.Count} 個");

        Assert.Pass($"✓ {successCount}/{testFiles.Count} 個のファイルでRound Tripテストが成功しました");
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB", "TB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}
