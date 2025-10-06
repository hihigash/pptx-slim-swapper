using System.CommandLine;
using PptxSlimSwapper.Services;

namespace PptxSlimSwapper;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // ルートコマンド
        var rootCommand = new RootCommand("PPTX Slim Swapper - PPTXファイル内のメディアを差し替えてファイルサイズを削減するツール");

        // outコマンド: メディアを外部に保存してプレースホルダーに差し替え
        var outCommand = CreateOutCommand();
        rootCommand.AddCommand(outCommand);

        // inコマンド: プレースホルダーを元のメディアに差し戻し
        var inCommand = CreateInCommand();
        rootCommand.AddCommand(inCommand);

        return await rootCommand.InvokeAsync(args);
    }

    private static Command CreateOutCommand()
    {
        var command = new Command("out", "PPTXファイル内のメディアを外部に保存し、プレースホルダーに差し替えます");

        var inputFileArgument = new Argument<FileInfo>(
            name: "input",
            description: "入力PPTXファイルのパス"
        ).ExistingOnly();

        var outputDirectoryOption = new Option<DirectoryInfo>(
            aliases: new[] { "--output", "-o" },
            description: "出力ディレクトリのパス(省略時はカレントディレクトリに'output'フォルダを作成)",
            getDefaultValue: () => new DirectoryInfo(Path.Combine(Directory.GetCurrentDirectory(), "output"))
        );

        command.AddArgument(inputFileArgument);
        command.AddOption(outputDirectoryOption);

        command.SetHandler(async (FileInfo inputFile, DirectoryInfo outputDirectory) =>
        {
            try
            {
                Console.WriteLine($"処理を開始します...");
                Console.WriteLine($"入力ファイル: {inputFile.FullName}");
                Console.WriteLine($"出力ディレクトリ: {outputDirectory.FullName}");
                Console.WriteLine();

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                var (outputPptxPath, manifestPath, mediaCount) = await PptxMediaSwapper.SwapOutMediaAsync(
                    inputFile.FullName,
                    outputDirectory.FullName
                );

                stopwatch.Stop();

                Console.WriteLine($"✓ 処理が完了しました! (所要時間: {stopwatch.ElapsedMilliseconds}ms)");
                Console.WriteLine();
                Console.WriteLine($"差し替えたメディア数: {mediaCount}");
                Console.WriteLine($"出力PPTXファイル: {outputPptxPath}");
                Console.WriteLine($"マニフェストファイル: {manifestPath}");
                Console.WriteLine();

                // ファイルサイズの比較
                var originalSize = new FileInfo(inputFile.FullName).Length;
                var newSize = new FileInfo(outputPptxPath).Length;
                var reduction = originalSize - newSize;
                var reductionPercent = (double)reduction / originalSize * 100;

                Console.WriteLine($"元のサイズ: {FormatFileSize(originalSize)}");
                Console.WriteLine($"新しいサイズ: {FormatFileSize(newSize)}");
                Console.WriteLine($"削減量: {FormatFileSize(reduction)} ({reductionPercent:F2}%)");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"エラー: {ex.Message}");
                Console.ResetColor();
                Environment.Exit(1);
            }
        }, inputFileArgument, outputDirectoryOption);

        return command;
    }

    private static Command CreateInCommand()
    {
        var command = new Command("in", "プレースホルダーを元のメディアに差し戻します");

        var inputFileArgument = new Argument<FileInfo>(
            name: "input",
            description: "差し替え済みPPTXファイルのパス"
        ).ExistingOnly();

        var manifestDirectoryOption = new Option<DirectoryInfo>(
            aliases: new[] { "--manifest-dir", "-m" },
            description: "マニフェストファイルがあるディレクトリのパス(省略時は入力ファイルと同じディレクトリ)",
            getDefaultValue: () => null!
        );

        command.AddArgument(inputFileArgument);
        command.AddOption(manifestDirectoryOption);

        command.SetHandler(async (FileInfo inputFile, DirectoryInfo? manifestDirectory) =>
        {
            try
            {
                // マニフェストディレクトリが指定されていない場合は入力ファイルと同じディレクトリを使用
                var manifestDir = manifestDirectory?.FullName ?? inputFile.DirectoryName ?? Directory.GetCurrentDirectory();

                Console.WriteLine($"処理を開始します...");
                Console.WriteLine($"入力ファイル: {inputFile.FullName}");
                Console.WriteLine($"マニフェストディレクトリ: {manifestDir}");
                Console.WriteLine();

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                var (outputPptxPath, restoredCount) = await PptxMediaSwapper.SwapInMediaAsync(
                    inputFile.FullName,
                    manifestDir
                );

                stopwatch.Stop();

                Console.WriteLine($"✓ 処理が完了しました! (所要時間: {stopwatch.ElapsedMilliseconds}ms)");
                Console.WriteLine();
                Console.WriteLine($"復元したメディア数: {restoredCount}");
                Console.WriteLine($"出力PPTXファイル: {outputPptxPath}");
                Console.WriteLine();

                // ファイルサイズの比較
                var slimSize = new FileInfo(inputFile.FullName).Length;
                var restoredSize = new FileInfo(outputPptxPath).Length;
                var increase = restoredSize - slimSize;

                Console.WriteLine($"差し替え済みサイズ: {FormatFileSize(slimSize)}");
                Console.WriteLine($"復元後のサイズ: {FormatFileSize(restoredSize)}");
                Console.WriteLine($"増加量: {FormatFileSize(increase)}");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"エラー: {ex.Message}");
                Console.ResetColor();
                Environment.Exit(1);
            }
        }, inputFileArgument, manifestDirectoryOption);

        return command;
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
