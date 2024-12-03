# 使い方

**実行環境**

- OS: Windows 11
- ランタイム: .NET Framework 4.8.1

**使用手順**

1. 出力先ディレクトリを設定
PowerShell スクリプト内で $outputDirectory を設定します。例として、デスクトップ上の "LocalizedHtml" フォルダに出力する場合:

'''powershell
$outputDirectory = Join-Path $basePath "LocalizedHtml"
'''

2. スクリプトを実行
PowerShell スクリプトを実行すると、指定した $outputDirectory にローカライズされた HTML ファイルが生成されます。

**出力先ディレクトリの指定方法**

1. PowerShell スクリプトを開く
  - ファイル名は resource_replacer.ps1 です。
2. $outputDirectory の値を変更
  - 好きな出力先フォルダのパスを指定してください。
```powershell
# 例: ドキュメントフォルダ内の "MyLocalizedHtml" フォルダに出力
$outputDirectory = "C:\Users\YourName\Documents\MyLocalizedHtml"
```
3. スクリプトを保存して実行
  - スクリプトを保存し、PowerShell から実行します。

**注意点**

出力先ディレクトリが存在しない場合、スクリプトが自動的にディレクトリを作成します。


# 本ツールのカスタム例

PowerShell スクリプトからHTML生成クラスを呼び出して英語のHTMLを出力する例

```powershell
# インスタンスの作成
$htmlGenerator = New-Object LocalizedHtmlGenerator

# パスの設定
$excelPath = "C:\Users\YourName\Documents\resources.xlsx"
$templatePath = "C:\Users\YourName\Documents\template.html"
$outputBaseName = "Content"
$cultureName = "English"

# HTML ファイルを生成
$htmlGenerator.GenerateHtml($outputBaseName, $cultureName, $templatePath, $excelPath)
```


# GenerateHtmlのファイルパスの設定方法と関連する変数についての解説

LocalizedHTMLGenerator.GenerateHtml() 解説

1. excelPath（Excel ファイルのパス）
  - 説明: リソースが記載された Excel ファイル（例: resources.xlsx）のファイルパスです。
  - 設定方法: 実際の Excel ファイルの場所をフルパスまたは相対パスで指定します。
  - 例: C:\Users\YourName\Documents\resources.xlsx
2. templatePath（テンプレート HTML ファイルのパス）
  - 説明: プレースホルダーが含まれたテンプレート HTML ファイルのパスです。プレースホルダーは {キー} の形式で記載します。
  - 設定方法: テンプレート HTML ファイルの場所を指定します。
  - 例: C:\Users\YourName\Documents\template.html
3. outputBaseName（出力ファイルのベース名）
  - 説明: 生成される HTML ファイルのベース名です。実際のファイル名は {outputBaseName}_{cultureName}.html となります。
  - 設定方法: 任意のベース名を指定します。
  - 例: "Content"
4. cultureName（カルチャ名）
  - 説明: 生成する言語の名前です。これは Excel ファイルのヘッダーに記載されている言語と一致する必要があります。
  - 設定方法: Excel ファイルのヘッダーにある言語名を文字列として指定します。
  - 例: "English", "Japanese", "French"


# 入力する Excel ファイル（翻訳文字列台帳）のフォーマットについての解説
Excel ファイルの構成

**ヘッダー行（1 行目）**: キーと各言語のカルチャ名を記載します。

- A1 セル: Key と記載します。
- B1 セル以降: 各言語の名前を記載します（例: English, Japanese, French）。

**データ行（2 行目以降）:**

- A 列（キー列）: プレースホルダーのキーを記載します（例: Title, Description）。
- B 列以降: 各言語に対応する翻訳テキストを記載します。

**具体的な例**

|A	|B	|C	|D|
|---|---|---|-|
|1	|Key	|English	|Japanese	|French |
|2	|Title	|Welcome	|ようこそ|	|Bienvenue  |
|3	|Description	|This is a sample.	|これはサンプルです。	|C'est un exemple.  |

**注意点**
キーはユニークである必要があります。テンプレート内のプレースホルダーと一致させます。
セルの内容はプログラム内でダブルクォーテーションで囲まれます。プログラム内でエスケープ処理を行っているためです。


# テンプレート HTML ファイルでのプレースホルダーの書き方

**テンプレート HTML ファイル（template.html）**

プレースホルダーの形式: {キー}

例: 
```html
<html>
<head>
    <title>{1}</title>
</head>
<body>
    <h1>{2}</h1>
    <p>{3}</p>
</body>
</html>
```

**置換の流れ**
プログラムはHTMLを1行毎に読み込み、各行について全てのキーとプレースホルダーをマッチング判定しています。
マッチした全てのプレースホルダー {キー} を Excel ファイルで指定した言語の翻訳テキストで置換します。
置換後にユーザーが指定したファイルパス設定をもとにHTMLファイルを出力します。


# 開発環境

- [.NET Framework 4.8 Developer](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)
- nuget (本プロジェクトに同梱)

**ビルド手順**

```cmd
.\env.bat
.\nuget_restore.bat
.\build.bat
```


# 本ツールのメンテナンス

**新しい翻訳対象の言語を追加する**

Excel ファイルのヘッダー行に新しいカルチャ名を追加し、対応する翻訳を記入します。

PowerShell スクリプトの $cultures 配列に新しいカルチャ名を追加します。

```powershell
$cultures = @("English", "Japanese", "French", "Spanish")
```

**ファイルパスの変更**

ファイルパスを変更する場合、PowerShell スクリプト内の対応する変数を更新してください。

- $basePath
- $templatePath
- $excelPath
- $outputDirectory

**CSVファイル**

プログラムは Excel ファイルを一時的に CSV ファイルに変換します（temp_resources.csv）。
CSV ファイルは同じディレクトリに生成されます。

**エラーハンドリング**

指定したカルチャ名が Excel ファイルに存在しない場合、エラーが発生します。
Excel ファイルが空の場合やフォーマットが正しくない場合もエラーになります。

**パフォーマンス**

大きなファイルを処理する場合、進行状況をコンソールに出力します（1000 行ごと）。
