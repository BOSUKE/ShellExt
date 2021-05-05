# ShellExt
Pathをいい感じにコピーしたりするWinシェル拡張

* ネットワークドライブをUNCに変換してクリップボードにコピー
* 複数のファイルを <ディレクトリ> とファイル名の行に分けてクリップボードにコピー  
  例:  
  <M:\VSProject\ShellExt\>  
  x64\  
  .gitattributes  
  .gitignore  
  dlldata.c  
  dllmain.cpp
* %USERPROFILE% の部分を削除してクリップボードにコピー  
  C:\Users\bosuke\OneDrive\ドキュメント\test.txt ⇒ <OneDrive\ドキュメント\test.txt>  
  (OneDrive や BoxDrive などのPathを他の人に伝えるのに使う）
* クリップボードにコピーされているファイルパスを開く
* 覚えたパスのショートカットを作成

## エクスプローラで右クリック
ファイルorフォルダを選択して**いる**状態で右クリック

![click_on_file](https://user-images.githubusercontent.com/19199772/117143308-a6a9d580-adeb-11eb-92c8-2b20f986a185.PNG)

ファイルorフォルダを選択して**いない**状態で右クリック

![click_on_directory](https://user-images.githubusercontent.com/19199772/117143319-a9a4c600-adeb-11eb-992c-b45244f8409a.PNG)

状態に応じて表示される項目は変化する。

## インストール/アンインストール
1. コマンドプロンプトを管理者として実行する。
2. ShellExt.dllを置いたディレクトリに移動する。
3. regsvr32 ShellExt.dll でインストール。regsvr32 /u ShellExt.dllでアンインストール。
