# 音声文字起こしアプリ 使用マニュアル

## 主な機能一覧
| 機能カテゴリ     | 内容                                                             |
|------------|----------------------------------------------------------------|
| 入力対応       | mp3 / m4a / wav / ogg / mp4 / avi / mov /wnv                                 |
| 自動変換       | 動画 / m4a → mp3 に変換（ffmpeg使用）        |
| ノイズ除去     | 音声ファイルを自動でノイズ除去（noisereduceライブラリ使用）                                   |
| 言語対応       | **自動言語検出、日本語・英語混在対応**（Whisperモデル small 使用）                                   |
| 分割処理       | 長時間ファイルを自動分割（30分ごと・カスタマイズ可）                                         |
| 出力形式       | txt / Word（docx） 選択可。txt / Word（docx）がデフォルト。Wordはスタイル整形済（フォントMeiryo UI・余白1inch・11pt） |
| 保存名自動     | 元ファイル名＋日付＋時刻で自動保存（同じフォルダ内）                                             |
| 改行処理       | 「。」や「.」で段落改行（日本語 / 英語 両対応）                                                 |
| 進捗表示       | 処理進捗バー（ファイル数＋分割単位反映）                                                       |
| 処理中断       | 処理中「中断」ボタンで停止可能                                                               |

---

## 使用方法（GUI操作手順）
### 1 ファイル読み込み   

### 2 出力形式チェック  
- txt（デフォルトON）  
- Word（デフォルトON）  

### 3 実行  
- **「文字起こし実行」ボタン**  
進捗バーが動きます。  
途中で止めたければ「中断」。

### 4 保存結果  
- 保存先：**元ファイルと同じフォルダ**
- ファイル名：`元ファイル名_YYYYMMDD_HHMMSS.txt` / `.docx`

---

## セットアップ方法
1. Pythonのインストールを行う。

公式サイト: python.org

2. 必要なライブラリのインストール
コマンドプロンプトやターミナルを開き、以下のコマンドを実行して、必要なライブラリをインストールする
pip install tkinterdnd2 whisper python-docx soundfile librosa noisereduce ffmpeg-python torch

3. ffmpegの準備
[ffmpeg公式サイト](https://ffmpeg.org/download.html) → ダウンロード

解凍して C:\ffmpeg\bin\ffmpeg.exe に配置

"ソースコード内：os.environ["PATH"] += os.pathsep + r"C:\ffmpeg\bin""

## 注意事項
### 1 音声が長時間（30分以上）の場合、自動分割 で処理

### 2 1ファイルずつ実施

### 2 変換時に 一時ファイル（_convert.mp3） を削除済

### 3 保存先：元ファイルと同じフォルダ に保存

### 4 空白や日本語ファイル名対応済

## 作者メモ
### Python / GUI初心者にも分かりやすい構成を意識

### 内部処理：Whisper + noisereduce + librosa

## ライセンス
- Whisper: OpenAI
- noisereduce: MIT
- tkinterdnd2 / librosa / soundfile: 各OSSライセンス
