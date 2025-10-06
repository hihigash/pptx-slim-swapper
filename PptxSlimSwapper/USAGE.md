# PPTX Slim Swapper - 使い方ガイド

## クイックスタート

### 1. ビルド

```powershell
cd c:\works\pptx-slim-swapper\PptxSlimSwapper
dotnet build -c Release
```

### 2. 実行方法

#### ヘルプの表示

```powershell
dotnet run -- --help
```

#### メディアを差し替える(out)

```powershell
# 基本的な使い方
dotnet run -- out <PPTXファイルパス>

# 例: presentation.pptx を処理
dotnet run -- out presentation.pptx

# 出力先を指定する場合
dotnet run -- out presentation.pptx --output ./my-output
```

**出力:**
- `output/presentation_slim.pptx` - サイズ削減されたPPTXファイル
- `output/swap-manifest.json` - 差し替え情報
- `output/media/*` - 元のメディアファイル

#### メディアを復元する(in)

```powershell
# 基本的な使い方
dotnet run -- in output/presentation_slim.pptx

# マニフェストディレクトリを指定する場合
dotnet run -- in output/presentation_slim.pptx --manifest-dir ./output
```

**出力:**
- `output/presentation_slim_restored.pptx` - 元のサイズに復元されたPPTXファイル

### 3. リリースビルドで実行

パフォーマンスを向上させるには、リリースモードでビルドして実行します:

```powershell
dotnet build -c Release
dotnet run -c Release -- out presentation.pptx
```

または、実行可能ファイルを直接実行:

```powershell
dotnet publish -c Release -o publish
.\publish\PptxSlimSwapper.exe out presentation.pptx
```

## コマンドラインオプション詳細

### `out` コマンド

```
dotnet run -- out <input> [options]

引数:
  <input>              入力PPTXファイルのパス

オプション:
  --output, -o <path>  出力ディレクトリのパス
                       (デフォルト: ./output)
  --help               ヘルプを表示
```

### `in` コマンド

```
dotnet run -- in <input> [options]

引数:
  <input>                    差し替え済みPPTXファイルのパス

オプション:
  --manifest-dir, -m <path>  マニフェストファイルがあるディレクトリ
                             (デフォルト: 入力ファイルと同じディレクトリ)
  --help                     ヘルプを表示
```

## トラブルシューティング

### エラー: "PPTXファイルが見つかりません"

- ファイルパスが正しいか確認してください
- 相対パスではなく絶対パスを試してください

### エラー: "マニフェストファイルが見つかりません"

- `in` コマンド実行時に、対応する `swap-manifest.json` が必要です
- `--manifest-dir` オプションで正しいディレクトリを指定してください

### 警告: "メディアパーツが見つかりません"

- PPTX ファイルが編集されて構造が変わった可能性があります
- 差し替え後のファイルは大きな構造変更を行わないことを推奨します

## 動作確認用のテストPPTX作成

PowerPointで簡単なプレゼンテーションを作成し、画像を数枚挿入してテストできます:

1. PowerPoint で新規プレゼンテーション作成
2. 画像をいくつか挿入（できるだけ大きいファイル）
3. `test.pptx` として保存
4. コマンドを実行:
   ```powershell
   dotnet run -- out test.pptx
   ```
5. サイズの削減を確認
6. 復元を実行:
   ```powershell
   dotnet run -- in output/test_slim.pptx
   ```
