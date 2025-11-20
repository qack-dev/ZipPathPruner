<#
.SYNOPSIS
    Zipファイル内のフォルダ構造を剪定（フラット化）するスクリプト。
    ZipPathPruner.ps1

.DESCRIPTION
    指定されたフォルダ内のZipファイルを再帰的に探索し、
    Zipファイル内のフォルダ構造を取り除き、ファイルをフラットに配置します。
    入れ子のZipファイルも同様に処理され、Zipの階層構造は維持しつつフォルダのみを除去します。

.PARAMETER SourcePath
    処理対象のZipファイルが含まれる入力フォルダのパス。

.PARAMETER DestPath
    剪定されたZipファイルを保存する出力フォルダのパス。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File ZipPathPruner.ps1 -SourcePath .\Input -DestPath .\Output
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$DestPath
)

# .NETアセンブリのロード
Add-Type -AssemblyName System.IO.Compression.FileSystem

# パスの解決と作成
if (-not (Test-Path $SourcePath)) {
    Write-Error "指定された入力フォルダが見つかりません: $SourcePath"
    return
}
$SourcePath = Resolve-Path $SourcePath
$DestPath = [System.IO.Path]::GetFullPath($DestPath)

# 一時フォルダのベース作成
$TempBase = Join-Path ([System.IO.Path]::GetTempPath()) "ZipPathPruner_$(Get-Random)"
New-Item -ItemType Directory -Path $TempBase -Force | Out-Null

# 剪定ロジック関数
function Prune-ZipStructure {
    param (
        [string]$CurrentZipPath,
        [string]$OutputFolder,
        [int]$Depth,
        [string]$TempBase
    )
    
    $extractDir = Join-Path $TempBase "Extract_$(Get-Random)"
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($CurrentZipPath, $extractDir)
    } catch {
        Write-Warning "Zipの解凍に失敗しました: $CurrentZipPath"
        return $Depth
    }

    $maxDepth = $Depth

    # ファイルを再帰的に取得 (フォルダは無視、ファイルのみ対象)
    $files = Get-ChildItem -Path $extractDir -Recurse -File

    foreach ($file in $files) {
        if ($file.Extension -eq ".zip") {
            # 入れ子Zipの処理
            
            # 1. 新しいZipの中身を入れる一時フォルダを作成
            $nestedZipContentDir = Join-Path $TempBase "NestedZipContent_$(Get-Random)"
            New-Item -ItemType Directory -Path $nestedZipContentDir -Force | Out-Null
            
            # 2. 再帰呼び出し (OutputFolderは新しいZipの中身用フォルダ)
            $nestedMaxDepth = Prune-ZipStructure -CurrentZipPath $file.FullName -OutputFolder $nestedZipContentDir -Depth ($Depth + 1) -TempBase $TempBase
            if ($nestedMaxDepth -gt $maxDepth) { $maxDepth = $nestedMaxDepth }
            
            # 3. 新しいZipを親のOutputFolderに作成
            $destNestedZip = Join-Path $OutputFolder $file.Name
            
            # 名前衝突回避
            $counter = 2
            while (Test-Path $destNestedZip) {
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $destNestedZip = Join-Path $OutputFolder "${nameWithoutExt}_${counter}.zip"
                $counter++
            }
            
            [System.IO.Compression.ZipFile]::CreateFromDirectory($nestedZipContentDir, $destNestedZip)
            
            # クリーンアップ
            Remove-Item -Path $nestedZipContentDir -Recurse -Force

        } else {
            # 通常ファイルの移動（フラット化）
            $fileName = $file.Name
            $destFile = Join-Path $OutputFolder $fileName
            
            # 名前衝突の回避
            $counter = 2
            while (Test-Path $destFile) {
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $ext = [System.IO.Path]::GetExtension($fileName)
                $destFile = Join-Path $OutputFolder "${nameWithoutExt}_${counter}${ext}"
                $counter++
            }
            
            Copy-Item -Path $file.FullName -Destination $destFile
        }
    }
    
    Remove-Item -Path $extractDir -Recurse -Force
    return $maxDepth
}

try {
    # 対象ファイルの収集
    $zipFiles = Get-ChildItem -Path $SourcePath -Filter "*.zip" -Recurse

    if ($zipFiles.Count -eq 0) {
        Write-Host "Zipファイルが見つかりませんでした。"
        return
    }

    # Zipファイルがフォルダ構造を持つか判定する関数
    function Test-ZipHasFolders {
        param (
            [string]$ZipPath
        )
        try {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
            $hasFolders = $false
            foreach ($entry in $archive.Entries) {
                if ($entry.FullName.IndexOf("/") -ge 0 -or $entry.FullName.IndexOf("\") -ge 0) {
                    $hasFolders = $true
                    break
                }
            }
            $archive.Dispose()
            return $hasFolders
        } catch {
            Write-Warning "Zipファイルの読み込みに失敗しました: $ZipPath"
            return $false
        }
    }

    # フォルダ構造チェック（確認用）
    $needsPruning = $false
    foreach ($zipFile in $zipFiles) {
        if (Test-ZipHasFolders -ZipPath $zipFile.FullName) {
            $needsPruning = $true
            break
        }
    }

    # ユーザー確認
    if ($needsPruning) {
        $confirm = Read-Host "フォルダ構造を持つZipファイルが見つかりました。フォルダがないよう剪定し $DestPath に作成しますか？(y/n)"
        if ($confirm -ne 'y') {
            Write-Host "処理を中止しました。"
            return
        }

        # 出力フォルダの作成（存在しない場合のみ）
        if (-not (Test-Path $DestPath)) {
            New-Item -ItemType Directory -Path $DestPath -Force | Out-Null
        }
    } else {
        Write-Host "剪定が必要なZipファイルは見つかりませんでした。処理を終了します。"
        return
    }

    # 相対パス取得用関数 (.NET Framework互換)
    function Get-RelativePath {
        param (
            [string]$BasePath,
            [string]$Target
        )
        $PathUri = [Uri]$Target
        $BaseUri = [Uri]($BasePath.TrimEnd('\') + '\')
        return [Uri]::UnescapeDataString($BaseUri.MakeRelativeUri($PathUri).ToString().Replace('/', '\'))
    }

    # メイン処理ループ
    $globalMaxDepth = 0
    foreach ($zipFile in $zipFiles) {
        # 出力先のパス計算（相対パスを維持）
        $relPath = Get-RelativePath -BasePath $SourcePath -Target $zipFile.FullName
        $destZipPath = Join-Path $DestPath $relPath
        $destZipDir = Split-Path $destZipPath -Parent
        
        if (-not (Test-Path $destZipDir)) {
            New-Item -ItemType Directory -Path $destZipDir -Force | Out-Null
        }

        # 既存の出力ファイルがあれば削除
        if (Test-Path $destZipPath) {
            Remove-Item -Path $destZipPath -Force
        }

        if (Test-ZipHasFolders -ZipPath $zipFile.FullName) {
            Write-Host "処理中 (剪定): $relPath"
            
            # 新しいZip用の一時フォルダ
            $newZipContentDir = Join-Path $TempBase "NewZip_$(Get-Random)"
            New-Item -ItemType Directory -Path $newZipContentDir -Force | Out-Null

            # 剪定実行
            $depth = Prune-ZipStructure -CurrentZipPath $zipFile.FullName -OutputFolder $newZipContentDir -Depth 1 -TempBase $TempBase
            if ($depth -gt $globalMaxDepth) {
                $globalMaxDepth = $depth
            }

            # 新しいZipの作成
            [System.IO.Compression.ZipFile]::CreateFromDirectory($newZipContentDir, $destZipPath)
            
            # クリーンアップ
            Remove-Item -Path $newZipContentDir -Recurse -Force
            
            Write-Host "完了: $relPath"
        } else {
            Write-Host "処理中 (コピー): $relPath"
            Copy-Item -Path $zipFile.FullName -Destination $destZipPath
            Write-Host "完了: $relPath"
        }
    }
    
    Write-Host "全ての処理が完了しました。"
    Write-Host "zipの階層は $globalMaxDepth です。"

} catch {
    Write-Error "予期せぬエラーが発生しました: $_"
} finally {
    # 全体の一時フォルダ削除
    if (Test-Path $TempBase) {
        Remove-Item -Path $TempBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

