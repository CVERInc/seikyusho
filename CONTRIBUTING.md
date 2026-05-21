# Contributing / 貢献ガイド

Issue、Pull Request、翻訳追加、すべて歓迎します 🙌

[**日本語**](#日本語) ・ [**English**](#english)

---

## 日本語

### 貢献の種類

| 種類 | 方法 |
|---|---|
| 🐛 バグ報告 | [Issue](https://github.com/CVERInc/seikyusho/issues) を作成、再現手順を記載 |
| 💡 機能提案 | Issue または Discussions で議論してから PR |
| 🌐 翻訳追加 | PR（[新言語追加ガイド](#-新言語の追加方法) を参照）|
| 📝 ドキュメント改善 | PR 直接 |
| 🎨 UI / UX 改善 | スクリーンショット付き Issue → PR |

### ローカル開発環境

```bash
# 1. リポジトリを clone
git clone https://github.com/CVERInc/seikyusho.git
cd seikyusho

# 2. clasp インストール（初回のみ）
npm install -g @google/clasp
clasp login

# 3. 自分の Apps Script プロジェクトに紐付け
cp .clasp.json.example .clasp.json
# .clasp.json の scriptId を自分の Apps Script ID に書き換える

# 4. push してテスト
clasp push -f
```

### 🌐 新言語の追加方法

例として「ドイツ語 (de-DE)」を追加する手順:

#### Step 1. `i18n.html` に翻訳辞書を追加

```javascript
// TRANSLATIONS オブジェクトに新しい locale を追加
'de-DE': {
  page_title: 'Rechnungserstellung',
  intro: 'Dieses Formular sendet eine Rechnungsanfrage an {{COMPANY_NAME}}...',
  // ... 既存の言語と同じ keys を全て翻訳
  // ... `{{COMPANY_NAME}}` placeholder は実行時に置換される
}
```

> 💡 既存の `ja-JP` / `en-US` / `zh-TW` / `es-ES` のいずれかを丸ごとコピーして翻訳すると漏れがない

#### Step 2. `index.html` に `<option>` と COMPANY_NAMES を追加

```html
<!-- lang-switcher 内 -->
<option value="de-DE">Deutsch</option>

<!-- COMPANY_NAMES 内 -->
'de-DE': <?!= JSON.stringify(companyNames['de-DE'] || '') ?>
```

#### Step 3. `Code.gs` の `doGet` を更新

```javascript
const companyNames = {
  // ...
  'de-DE': settings['company_name_de-DE'] || ''  // 追加
};
```

#### Step 4. `Code.gs` の `validatePayload_` を更新

```javascript
language: ['ja-JP', 'en-US', 'zh-TW', 'es-ES', 'de-DE'].indexOf(p.language) >= 0
  ? p.language : 'ja-JP',
```

#### Step 5. `Setup.gs` の `setupSettingsSheet_` に追加

```javascript
['company_name_de-DE', 'YOUR_COMPANY_NAME_DE_DE', 'Empfängername (Deutsch)'],
```

#### Step 6. `Setup.gs` の `addMissingSettings` に追加

既存ユーザーが新言語にアップグレードできるように、`allDefaults` 配列に同じ row を追加（value は空欄に）:

```javascript
['company_name_de-DE', '', 'Empfängername (Deutsch)'],
```

#### Step 7. （その言語圏の主要通貨を追加）

新言語のユーザーが使う通貨も合わせて追加すると親切です。例: ドイツなら EUR は既存なので不要、韓国語 ko-KR を追加するなら KRW を追加する。

通貨追加の手順:
- `index.html` の `<option value="XXX">XXX (記号)</option>`
- `Code.gs` の `validatePayload_` 配列
- `Code.gs` の `formatCurrency_` の if 分岐
- `script.html` の前端 formatCurrency

#### Step 8. README を更新

- `[i18n badge]` に新 locale を追加
- 設定項目表に `company_name_xx-XX` を追加
- 多言語対応の説明文を更新

### 💴 新通貨の追加方法

| 場所 | 追加内容 |
|---|---|
| `index.html` `<select name="currency">` | `<option value="XXX">XXX (記号)</option>` |
| `Code.gs` `validatePayload_` | enum 配列に `'XXX'` を追加 |
| `Code.gs` `formatCurrency_` | `if (currency === 'XXX') return '記号' + formatted;` |
| `script.html` 前端 formatCurrency | 同上 |
| `Setup.gs` `populateUsageSheet` 入力項目欄 | `通貨(JPY/TWD/USD/EUR/XXX)` |
| `README.md` Multi-currency 説明 | 通貨を追加 |

### コードスタイル

- インデント: スペース 2 つ
- セミコロン: 必須
- 文字列: シングルクォート優先（ただし日本語テンプレートでは可読性を優先）
- 関数名: `camelCase`、private 関数は末尾に `_`（例: `regenerateInvoice_`）
- 定数: `UPPER_SNAKE_CASE`（例: `WITHHOLDING_THRESHOLD`）
- placeholder: `{{ALL_CAPS}}` 形式（例: `{{INVOICE_NO}}`）
- BCP 47 locale code: ハイフン区切り（例: `ja-JP`, `en-US`）

### Pull Request の心得

- 1 PR = 1 機能 / 1 修正（混ぜない）
- コミットメッセージは短く明確に（日本語/英語どちらでも OK）
- 大きな変更前に Issue で議論を
- スクリーンショット歓迎（UI 変更時は必須）

### 質問・相談

- 技術的な質問: [Discussions](https://github.com/CVERInc/seikyusho/discussions)
- セキュリティに関する報告: 公開しない (private email を使用してください)

---

## English

### Types of Contributions

| Type | How |
|---|---|
| 🐛 Bug reports | Create an [Issue](https://github.com/CVERInc/seikyusho/issues) with reproduction steps |
| 💡 Feature requests | Discuss in Issue or Discussions before PR |
| 🌐 Translations | PR (see [Adding a new language](#-adding-a-new-language)) |
| 📝 Documentation | PR directly |
| 🎨 UI/UX improvements | Issue with screenshots → PR |

### Local Development Setup

```bash
git clone https://github.com/CVERInc/seikyusho.git
cd seikyusho

npm install -g @google/clasp
clasp login

cp .clasp.json.example .clasp.json
# Edit .clasp.json: set scriptId to your own Apps Script project ID

clasp push -f
```

### 🌐 Adding a new language

Example: adding German (`de-DE`):

#### Step 1. Add translation dictionary in `i18n.html`

```javascript
'de-DE': {
  page_title: 'Rechnungserstellung',
  intro: 'Dieses Formular sendet eine Rechnungsanfrage an {{COMPANY_NAME}}...',
  // ... translate all keys
  // ... `{{COMPANY_NAME}}` is replaced at runtime
}
```

> 💡 Tip: copy an existing language block (e.g. `en-US`) entirely and translate, to avoid missing keys.

#### Step 2. Add `<option>` and COMPANY_NAMES key in `index.html`

```html
<option value="de-DE">Deutsch</option>

'de-DE': <?!= JSON.stringify(companyNames['de-DE'] || '') ?>
```

#### Step 3. Update `Code.gs` `doGet`

```javascript
const companyNames = {
  // ...
  'de-DE': settings['company_name_de-DE'] || ''
};
```

#### Step 4. Update `Code.gs` `validatePayload_`

```javascript
language: ['ja-JP', 'en-US', 'zh-TW', 'es-ES', 'de-DE'].indexOf(p.language) >= 0
  ? p.language : 'ja-JP',
```

#### Step 5. Update `Setup.gs` `setupSettingsSheet_`

```javascript
['company_name_de-DE', 'YOUR_COMPANY_NAME_DE_DE', 'Empfängername (Deutsch)'],
```

#### Step 6. Update `Setup.gs` `addMissingSettings`

For existing users to upgrade smoothly, add the same key with empty value to `allDefaults`:

```javascript
['company_name_de-DE', '', 'Empfängername (Deutsch)'],
```

#### Step 7. (Add the relevant currency if needed)

For the new language's primary currency. E.g., German → EUR (already exists). Korean → add KRW. See [Adding a new currency](#-adding-a-new-currency).

#### Step 8. Update README

- Add to the i18n badge
- Add `company_name_xx-XX` to the configuration table
- Update the multilingual description

### 💴 Adding a new currency

| Where | What to add |
|---|---|
| `index.html` `<select name="currency">` | `<option value="XXX">XXX (symbol)</option>` |
| `Code.gs` `validatePayload_` | Add `'XXX'` to enum array |
| `Code.gs` `formatCurrency_` | `if (currency === 'XXX') return 'symbol' + formatted;` |
| `script.html` front-end formatCurrency | Same |
| `Setup.gs` `populateUsageSheet` input items field | `通貨(JPY/TWD/USD/EUR/XXX)` |
| `README.md` Multi-currency description | Add the currency |

### Code Style

- Indentation: 2 spaces
- Semicolons: required
- Strings: prefer single quotes (Japanese template literals OK for readability)
- Function names: `camelCase`, private functions end with `_` (e.g. `regenerateInvoice_`)
- Constants: `UPPER_SNAKE_CASE` (e.g. `WITHHOLDING_THRESHOLD`)
- Placeholders: `{{ALL_CAPS}}` (e.g. `{{INVOICE_NO}}`)
- BCP 47 locale codes: hyphen-separated (e.g. `ja-JP`, `en-US`)

### Pull Request Etiquette

- 1 PR = 1 feature / 1 fix (don't mix concerns)
- Keep commit messages short and clear (Japanese or English OK)
- Discuss large changes in an Issue first
- Screenshots welcome (required for UI changes)

### Questions

- Technical questions: [Discussions](https://github.com/CVERInc/seikyusho/discussions)
- Security issues: please do NOT open a public issue — use private email

---

Thanks for considering contributing! 🙌
