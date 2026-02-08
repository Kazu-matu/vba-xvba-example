---

## 推奨運用と目的

### ソース抽出と管理の流れ

- VBA（Excel）からのソースコード抽出（エクスポート）は「最初のみ」行います。
- 抽出したソースは、`/src`フォルダ配下で用途ごとにフォルダ分け（例：`Module/`, `Class/`, `Form/`など）して管理します。
- ソースコードの編集・デバッグは必ずVS Code上で行い、編集後はExcelに取り込んで実行します（Excel側からの再エクスポートは推奨しません）。
- この運用により、ソース管理が一方通行（VS Code→Excel）となり、コードのメンテナンス性・品質が向上します。

### 目的

- ソースコードを用途別にフォルダ分けすることで、管理・メンテナンスが容易になります。
- フォルダやファイル名は、フレームワーク（例：Nexus-VBA等）のルールに従って整理してください。

---
# XVBA × Git 運用マニュアル（初心者向け）

---

## このマニュアルの目的

ExcelのVBA開発を、より安全・便利にチームで進めるための方法をまとめています。
XVBAとGitを使うことで、VBAコードをバージョン管理し、複数人での開発やバックアップ、過去の変更履歴の確認が簡単になります。

ただし、ExcelのVBAは「Shift-JIS」という日本語の文字コードを使うため、そのままGitで管理すると文字化けが起きやすいです。
このマニュアルでは、文字化けを防ぎ、快適にVBA開発を進めるための設定方法をやさしく解説します。

---

## こんな人におすすめ

- ExcelのVBAを複数人で安全に管理したい
- 文字化けせずにGitHubでVBAコードを共有したい
- バックアップや過去の変更履歴を残したい

---

## 用語解説（初心者向け）

- **VBA**: Excelなどで使えるプログラミング言語。マクロの中身。
- **Git**: ファイルの変更履歴を記録・管理できるツール。
- **リポジトリ**: Gitで管理するファイルの集まり。
- **エンコーディング**: 文字をパソコンで保存する際のルール。日本語は「Shift-JIS」や「UTF-8」などがある。
- **インデックス**: Gitで「次に記録するよ」と一時的に覚えておく場所。

---

## 図解：文字コード変換の流れ

```
Excel (Shift-JIS)
    ↑   ↓
VS Code (Shift-JIS)
    ↑   ↓
Git（コミット時にUTF-8へ自動変換）
    ↑   ↓
GitHub（UTF-8で保存）
```

---

## 困ったときは？（トラブルシューティング）

「コミットしたのに文字化けする」「VS Codeで日本語が変になる」など、よくある質問は[トラブルシューティング](#トラブルシューティング)の章を参照してください。

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
# VBA関連ファイルをShift-JISとして扱い、Git上ではUTF-8で管理する
*.bas text working-tree-encoding=shift_jis
*.cls text working-tree-encoding=shift_jis
*.frm text working-tree-encoding=shift_jis
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

# 5. 補足: .gitattributesの注意点・トラブルシューティング

## .gitattributesの注意点

- `working-tree-encoding`属性は **Git 2.21以上** でサポートされています。バージョンが古い場合は `git --version` で確認してください。
- 既存ファイルのエンコーディングを変更した場合、以下の手順でGitに再登録してください。
    1. ファイルをShift-JISで保存し直す
    2. インデックスから外す: `git rm --cached <ファイル名>`
    3. 再度追加: `git add <ファイル名>`
    4. コミット: `git commit -m "Re-encode <ファイル名> as Shift-JIS"`

# 6. XVBAのconfig.json記載例と内容説明

XVBAでは、`config.json`（または`xvba.json`）でExcelファイルやVBAフォルダのパスなどを指定します。

### config.json の記載例
```json
{
    "excel_file": "Test.xlsm",
    "vba_folder": "vba-files",
    "filename_format": "name_only",
    "export_on_save": true,
    "import_on_start": false
}
```

#### 各項目の説明
- **excel_file**: 対象となるExcelファイル名またはパス。
- **vba_folder**: VBAソースファイル（.bas, .cls, .frm等）を格納するフォルダ。
- **filename_format**: ファイル名の形式（例: `name_only`はシート名やモジュール名のみ）。
- **export_on_save**: trueの場合、VBAコードを保存時に自動でエクスポート。
- **import_on_start**: trueの場合、XVBA起動時にVBAコードを自動でインポート。

---

# 7. working-tree-encodingの補足説明

`.gitattributes`の`working-tree-encoding`属性は、Git 2.21以降で利用可能な機能です。
この属性を使うことで、リポジトリ内（Git管理下）は常にUTF-8で統一し、作業ツリー（ローカルファイル）は指定したエンコーディング（例: Shift-JIS）で自動的に変換されます。

## 公式ドキュメント要約
> working-tree-encoding属性は、ファイルをチェックアウトする際に指定したエンコーディング（例: shift_jis）で作業ツリーに展開し、コミット時にはUTF-8に自動変換します。これにより、異なるエンコーディングを必要とするツールや環境でも、Gitリポジトリ内は常にUTF-8で一貫性を保てます。

詳細は[Git公式ドキュメント: working-tree-encoding](https://git-scm.com/docs/gitattributes#_working_tree_encoding)を参照してください。
---




## トラブルシューティング

### Q. Gitでコミットしたのに日本語が文字化けする
- `.gitattributes`の設定が正しいか確認
- Gitのバージョンが2.21以上か確認（`git --version`）
- 既存ファイルは一度 `git rm --cached <ファイル>` でインデックスから外し、再度 `git add` してください

### Q. コミット内容の文字化け確認方法
```sh
git show HEAD:vba-files/Module/Mod1.bas
```
- 上記コマンドで日本語が正しく表示されていればOKです

### Q. VS Codeで文字化けする場合
- ファイルを開いた状態で「エンコーディング付きで保存」→「Shift JIS」を選択してください。

## XVBA拡張機能のインストール方法

1. VS Code左側の拡張機能アイコンをクリック
2. 検索ボックスに「XVBA」と入力し、拡張機能をインストール

## .gitattributesの例（バイナリファイルも含む）
```gitattributes
*.bas text working-tree-encoding=shift_jis
*.cls text working-tree-encoding=shift_jis
*.frm text working-tree-encoding=shift_jis
*.frx binary
```

## 参考リンク

- [XVBA公式リポジトリ](https://github.com/xvba/xvba)
- [Git公式ドキュメント: working-tree-encoding](https://git-scm.com/docs/gitattributes#_working_tree_encoding)
XVBA（Live Server VBA）は、VS CodeとExcelをリアルタイムで同期できる非常に強力な拡張機能です。ただし、XVBA単体では「UTF-8とShift-JISの変換」を自動で行う機能が弱いため、**Gitの設定（.gitattributes）**と組み合わせるのが2026年現在のベストプラクティスです。

---
