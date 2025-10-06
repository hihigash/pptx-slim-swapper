# PPTX Slim Swapper

PowerPointファイル(PPTX)内の画像や動画を一時的に外部保存し、プレースホルダーに置き換えることでファイルサイズを削減するコマンドラインツールです。

## 特徴

- 📉 **ファイルサイズの大幅削減**: 画像・動画を小さなプレースホルダーに置き換え
- 🔄 **完全な復元機能**: 後から元のメディアファイルに差し戻し可能
- 🔒 **安全な処理**: 元のファイルは保持し、新しいファイルを生成
- 🎯 **簡単な操作**: シンプルなコマンドラインインターフェース

## 技術仕様

- **.NET 8** で実装
- **OpenXML SDK** を使用した静的なファイル解析と編集
- **System.CommandLine** によるモダンなCLI実装
- プレースホルダーにはメディア識別情報を埋め込み

## インストール

```bash
dotnet build -c Release
```

## 使用方法

### 1. メディアを外部保存(out)

PPTXファイル内の画像・動画を外部に保存し、プレースホルダーに差し替えます。

```bash
dotnet run -- out <入力PPTXファイル> [オプション]
```

#### 例

```bash
# 基本的な使い方
dotnet run -- out presentation.pptx

# 出力先を指定
dotnet run -- out presentation.pptx --output ./my-output
dotnet run -- out presentation.pptx -o ./my-output
```

#### 出力

- `<出力ディレクトリ>/<ファイル名>_slim.pptx`: 差し替え済みPPTXファイル
- `<出力ディレクトリ>/swap-manifest.json`: 差し替え情報を記録したマニフェストファイル
- `<出力ディレクトリ>/media/`: 抽出されたメディアファイル

### 2. メディアを復元(in)

プレースホルダーを元のメディアファイルに差し戻します。

```bash
dotnet run -- in <差し替え済みPPTXファイル> [オプション]
```

#### 例

```bash
# 基本的な使い方(マニフェストが同じディレクトリにある場合)
dotnet run -- in presentation_slim.pptx

# マニフェストディレクトリを明示的に指定
dotnet run -- in presentation_slim.pptx --manifest-dir ./output
dotnet run -- in presentation_slim.pptx -m ./output
```

#### 出力

- `<入力ファイルと同じディレクトリ>/<ファイル名>_restored.pptx`: 復元されたPPTXファイル

## ワークフロー例

```bash
# 1. 大きなPPTXファイルを差し替え
dotnet run -- out large-presentation.pptx -o ./output

# 出力:
#   ./output/large-presentation_slim.pptx (サイズ削減済み)
#   ./output/swap-manifest.json
#   ./output/media/* (元のメディアファイル)

# 2. 差し替え済みファイルで作業(編集、共有など)
# ...

# 3. 必要に応じて元に戻す
dotnet run -- in ./output/large-presentation_slim.pptx -m ./output

# 出力:
#   ./output/large-presentation_slim_restored.pptx (元のサイズに復元)
```

## プロジェクト構成

```
PptxSlimSwapper/
├── Models/
│   ├── MediaInfo.cs          # メディア情報モデル
│   └── SwapManifest.cs       # マニフェストモデル
├── Services/
│   ├── PlaceholderGenerator.cs  # プレースホルダー生成
│   └── PptxMediaSwapper.cs      # PPTX処理メインロジック
├── Program.cs                # エントリーポイント
└── PptxSlimSwapper.csproj
```

## 動作原理

### 差し替え処理(out)

1. PPTXファイルを読み込み
2. 画像パート(ImagePart)と動画パート(VideoPart)を抽出
3. 各メディアに一意なIDを割り当て
4. メディアデータを外部ファイルとして保存
5. 小さなプレースホルダー画像を生成してPPTX内に差し替え
6. マニフェストファイルに差し替え情報を記録

### 復元処理(in)

1. マニフェストファイルを読み込み
2. 差し替え済みPPTXファイルを開く
3. マニフェストの情報に基づいて各メディアパートを検索
4. 保存されていた元のメディアデータで上書き
5. 復元済みPPTXファイルとして保存

## 必要な依存関係

- DocumentFormat.OpenXml 3.2.0
- System.CommandLine 2.0.0-beta4.22272.1
- System.Drawing.Common 9.0.9
