XVBA（Live Server VBA）は、VS CodeとExcelをリアルタイムで同期できる非常に強力な拡張機能です。ただし、XVBA単体では「UTF-8とShift-JISの変換」を自動で行う機能が弱いため、**Gitの設定（.gitattributes）**と組み合わせるのが2026年現在のベストプラクティスです。

---

## 1. XVBA の基本セットアップ

まずは、プロジェクトのルートディレクトリに `xvba.json` を作成し、Excelファイルとソースコードの紐付けを行います。

### `xvba.json` の例

```json
{
    "excel_file": "bin/ChemicalManager.xlsm",
    "vba_folder": "src",
    "filename_format": "name_only"
}

```

* **excel_file**: 対象のExcelブックへのパス。
* **vba_folder**: VBAソースを書き出すフォルダ。

---

## 2. 【重要】文字コード変換の自動化設定

「GitHub上はUTF-8、ローカル（VBA）はShift-JIS」を実現するために、Gitの機能を活用します。プロジェクトルートに `.gitattributes` というファイルを作成してください。

### `.gitattributes` の設定

この設定をすることで、**Gitがコミット時に自動でUTF-8に変換し、チェックアウト時に自動でShift-JISに戻してくれます。**

```text

```

### VS Code 側の設定 (`.vscode/settings.json`)

VS Codeのエディタ上でも正しく表示されるよう、拡張子ごとにエンコードを指定します。

```json
{
    "files.associations": {
        "*.bas": "vb",
        "*.cls": "vb",
        "*.frm": "vb"
    },
    "[vb]": {
        "files.encoding": "shiftjis"
    }
}

```

---

## 3. 実運用ワークフロー（Excel ⇔ GitHub）

設定が完了したら、以下の手順で開発を進めます。

### 手順 A：Excel からコードを抽出する (Export)

1. Excelファイルを開きます。
2. VS Codeで `Ctrl + Shift + P` を押し、`XVBA: Export VBA` を実行します。
3. `src` フォルダに `.bas` や `.cls` が書き出されます（この時点では **Shift-JIS**）。

### 手順 B：GitHub へプッシュする (UTF-8化)

1. VS Codeのソース管理（Git）で変更をステージングします。
2. このとき、`.gitattributes` の効果で、**GitHubのリポジトリ内では自動的に UTF-8** に変換されて保存されます。
3. GitHub上のプルリクエストでは、日本語のコメントも文字化けせずにレビュー可能です。

### 手順 C：GitHub から取り込み、Excelへ反映 (Import)

1. `git pull` を実行します。このとき、ローカルには **Shift-JIS** としてファイルが展開されます。
2. VS Codeで `XVBA: Run Live Server` を起動します。
3. VS Codeでコードを編集して保存すると、**リアルタイムでExcel内のVBAコードが更新されます。**

---

## 4. XVBA を使う上での注意点

* **バイナリファイルの競合:** フォーム（`.frm`）に付随する `.frx` ファイルはバイナリです。これらは `.gitattributes` の対象外（binary指定）にし、コンフリクトが起きないよう「マクロ付きブック本体」とは別にソースコード主導で管理するのがコツです。
* **信頼センターの設定:** 他のツール同様、Excelの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックが入っているか確認してください。

---

### 次のステップへの提案

このXVBAの構成に、以前お話しされていた「Nexus-VBA」フレームワークの標準ディレクトリ構造（`App/`, `Domain/`, `Infrastructure/` など）を自動生成するタスクを追加してみるのはいかがでしょうか？

もしよろしければ、**VS Codeから「Nexus-VBA」の雛形フォルダを一括作成するスクリプト（PowerShell）**を作成しましょうか？