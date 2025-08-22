# ChangeLens - Revit モデル差分検出アドイン

## 概要
**ChangeLens** は、Autodesk Revit 用の差分検出アドインです。  
古い Revit モデルと現在のモデルを比較し、**追加・変更・削除された要素**や**パラメータ変更**を視覚的にハイライトし、Excel に出力します。  
主に設備系（配管、ダクト、ケーブルラック）の設計変更確認やレビューに最適化されています。

---

## 主な機能
- **要素の差分検出**
  - 追加: 赤色でハイライト
  - 変更: 青色でハイライト
  - パラメータ変更: オレンジでハイライト
  - 削除: Excel に記録のみ（モデル上には表示されません）
- **進捗表示**
  - WPF の `ProgressBar` に差分処理の進捗をリアルタイム表示
- **Excel 出力**
  - OpenXML を使って差分結果を Excel (`DiffReport.xlsx`) に書き出し
  - 出力内容: `Id`, `Category`, `Name`, `Status`
- **安全な Revit API 呼び出し**
  - 非同期処理は `ExternalEvent` を使用して Revit API の制約に対応

---

## 対象カテゴリ
- PipeCurves（配管）
- DuctCurves（ダクト）
- CableTray（ケーブルトレイ）

---

## UI
- **ファイル選択**: 古いモデル（.rvt）を選択
- **開始ボタン**: 差分処理を実行
- **進捗バー**: 処理の進捗を表示
- Excel 出力はデスクトップに固定されます

---

## 技術的ポイント
- **DiffProcessor**: 差分検出ロジックを担当
  - 要素の位置比較（LocationPoint, LocationCurve）
  - パラメータ比較
  - 色付けと Excel 出力
- **DiffHandler + ExternalEvent**: 非同期処理を安全に Revit API 内で実行
- **XYZExtensions**: 座標比較用の拡張メソッド `IsAlmostEqualTo`
- **WPF ダイアログ**: `DiffDialog.xaml` により簡単操作可能
- **Revit Command**: `Command` クラスで Revit メニューから呼び出し

---

## 使い方
1. Revit でアドインをロード
2. メニューから **ChangeLens** を実行
3. Diff Dialog が表示される
4. 「参照」ボタンで **古いモデル** を選択
5. 「開始」を押すと差分処理が開始され、進捗バーに反映
6. デスクトップに `DiffReport.xlsx` が出力される
7. 差分は Revit モデル上で色分けされ確認可能

---

## インストール方法
1. 本リポジトリをクローン
2. Visual Studio でプロジェクトをビルド
3. 出力 DLL を Revit の AddIns フォルダに配置
4. Revit を再起動してメニューから実行

---

## 注意事項
- 非同期処理のため、Revit の UI スレッド以外から直接モデル操作はできません。`ExternalEvent` を利用しています。
- Excel 出力先はデスクトップに固定。必要に応じてコード内でパスを変更可能です。
- 対象カテゴリ以外の要素は無視されます。

---

## 今後の拡張案
- 任意のカテゴリ追加対応
- Excel 出力先の変更機能
- 差分ハイライト色のカスタマイズ
- 複数フロアや複数モデルの一括比較

## Qiitaリンク
https://qiita.com/KengoTanaka-BIM/items/0a54b51007f6b590fcc5
