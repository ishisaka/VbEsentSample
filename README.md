VbEsentSample
=============

これはVBでESENT Managed Interfaceライブラリを使用してESEデータベースを使用するためのサンプルプロジェクトです。

## ビルド環境

Visual Studio 2013 (Pro以上を推奨)

## ESEについて

ESE(Extensible Storage Engine)は拡張されたISAM（Indexed and Sequential Access Method）エンジンです。ESEは、アプリケーションに組み込まれることを前提としていて、アプリケーションからテーブルの定義、定義したテーブルに対する追加、変更、削除の操作、保存したデータに対するインデックスを使用した検索、シーケンシャルなカーソル操作が可能になっています。

ESEはWindows 2000から標準機能(API)として組み込まれています。

ESEはWindowsのActive DirectoryのストレージやExchange Serverのストレージとして使用されています。

ESEのデータフォーマットはJETデータベースに準じた物となっていますが、Access 97のMDBファイルとは直接的な互換性は持っていません。また、MDBでは得ることが不可能なロギングと、そこからのリカバリーの機能があり、耐障害性が強化されています。

ADやExchangeで利用されている事からわかるように、高いパフォーマンスも有しています。


[Extensible Storage Engine(英文)](http://msdn.microsoft.com/en-us/library/windows/desktop/gg269259(v=exchg.10).aspx)

##ESENT Managed Interface(ManagedEsent)ついて

ManagedEsentは.NETのコードからESEのAPIを使用できるようにするための互換ライブラリで、Windows XP以降で動作するように作成されています。

<https://managedesent.codeplex.com/>
