# ZipPathPruner

ZipPathPrunerは、Zipファイル内のフォルダ構造を自動的に「剪定（フラット化）」するPowerShellスクリプトです。
複雑に入れ子になったフォルダを取り除き、ファイルや入れ子のZipファイルを直感的にアクセスしやすい構造に整理します。

※このリポジトリは、Google AntigravityによるVibe codingで作成されました。

## 特徴

*   **フォルダの除去**: Zip内のフォルダ階層をなくし、ファイルをフラットに配置します。
*   **入れ子Zip対応**: Zipの中にZipがある場合も、その構造を維持しつつ、中身のフォルダのみを除去します。
*   **再帰的処理**: 指定したフォルダ内のすべてのZipファイルを対象にします。
*   **安全設計**:
    *   元ファイルは変更せず、指定した出力フォルダに新しいZipを作成します。
    *   ファイル名が重複する場合は、自動的に連番を付与して上書きを防ぎます。

## 前提条件

*   Windows 10 または Windows 11
*   PowerShell 5.1 以上（標準でインストールされています）

## 使い方

1.  `ZipPathPruner.ps1` を適当なフォルダに保存します。
2.  そのフォルダを右クリックし、コマンドプロンプト（またはPowerShell）を開きます。
3.  以下のコマンドを実行します。

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -File ZipPathPruner.ps1 -SourcePath "入力フォルダのパス" -DestPath "出力フォルダのパス"
```

### 実行例

`C:\Data\Input` にあるZipファイルを処理して、`C:\Data\Output` に保存する場合:

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -File ZipPathPruner.ps1 -SourcePath C:\Data\Input -DestPath C:\Data\Output
```

実行すると、フォルダ構造を持つZipファイルが見つかった場合に確認メッセージが表示されます。
`y` を入力しEnterを押すと処理が開始されます。

## 剪定（Pruning）のイメージ

このスクリプトは、Zipファイルの中身を以下のように変更します。

**処理前（Before）**:
```text
MyArchive.zip
 └─ folder1
     ├─ document.txt
     └─ subfolder
         ├─ image.png
         └─ nested.zip
             └─ deep_folder
                 └─ data.csv
```

**処理後（After）**:
```text
MyArchive.zip
 ├─ document.txt
 ├─ image.png
 └─ nested.zip
     └─ data.csv
```

*   `folder1` や `subfolder` などのコンテナ（フォルダ）が取り除かれます。
*   中身のファイル（`document.txt`, `image.png`）はZipのルートに移動します。
*   `nested.zip` もルートに移動し、さらにその中身も同様に剪定されます（`deep_folder` が消え `data.csv` が `nested.zip` のルートへ）。

## 注意事項

*   同名のファイルが異なるフォルダにある場合（例: `folderA/test.txt` と `folderB/test.txt`）、処理後は `test.txt` と `test_2.txt` のようにリネームされて保存されます。
