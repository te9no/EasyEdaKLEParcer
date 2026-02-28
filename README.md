# KLE Importer Pro (EasyEDA Pro Extension)

KLE (Keyboard Layout Editor) の JSON/TXT を読み込み、**KLEのレジェンド内に書いた `SWxx` に従って** EasyEDA Pro のPCB上のフットプリント `SWxx`（任意で `Dxx`）を対応する位置へ移動します。

## できること

- KLEの各キーのレジェンド文字列から `SW1` / `SW2` / ... を検出
- PCB上の同名デジグネータ `SW1` / `SW2` / ... を、そのキー位置へ配置
- 任意でダイオード `D1` / `D2` / ... も、指定オフセット (mm) で配置

## 前提

- EasyEDA Pro (PCBエディタ) で実行してください
- PCB上に `SW1` `SW2` ...（必要なら `D1` `D2` ...）が存在すること
- KLE側のキーのレジェンド文字列に `SWxx` を含めてください（例: `"#\\n3\\n...\\nSW4"` のように、どこかに `SW4` が入っていればOK）

## インストール

1. リリースされた `.eext` を用意（例: `kle-importer-pro_v0.1.20.eext`）
2. EasyEDA Pro → `Extensions Manager` → `Import Extensions` → `.eext` を選択

## 使い方

1. EasyEDA Pro の PCB エディタを開く
2. 上部メニューに `KLE Importer Pro` が出るので `Open...` を押す
3. `JSON/TXT を選択` → KLEファイルを選ぶ
4. ピッチ、ダイオード配置/オフセットを設定
5. `実行`

### SW番号の付け方（重要）

- KLEの各キーのレジェンド文字列に `SW<number>` を含めてください
- 例: `SW4` が書かれているキーは、PCBの `SW4` をその位置へ移動します
- `SWxx` が書かれていないキーは **無視** されます

## トラブルシュート

### `レジェンドに SWxx が見つかりませんでした`
- KLEのレジェンドに `SW1` のような文字列が入っているか確認してください

### `PCB API が利用できません`
- PCBエディタを開いている状態で実行してください（回路図だけの状態では動きません）

### `Invalid or unexpected token`
- KLEファイルがJSONとして壊れている/余計な前置きがある可能性があります
- 本拡張はBOM除去や配列部分抽出も試しますが、それでもダメならエラーダイアログに出る `preview:` を見て原因を確認してください

## 開発・ビルド

```powershell
npm ci
npm run build
```

ビルド成果物は `build/dist/kle-importer-pro_v<version>.eext` に出力されます（versionは `extension.json` に従います）。

## リリース手順（GitHub）

1. `extension.json` の `version` を更新
2. `package.json` / `package-lock.json` の `version` も同じ値に更新
3. `npm ci` → `npm run build`
4. GitHub Releases でタグを切って `.eext` を添付

## ライセンス

Apache-2.0（`LICENSE` を参照）
