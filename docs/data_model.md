# 共通売上データモデル 設計書

- 状態: **MVP-1 で実装済み**（Excel 請求書・`template_a` の範囲）。MVP-2 で架空の内税・源泉徴収型書式 `template_b` を追加済み。MVP-2 以降の範囲は設計・実装継続中
- 対象スキーマバージョン: `0.1`
- 本書に記載するデータ例は、すべて **完全な架空データ** です。

## 1. 目的と範囲

Excel 請求書、PDF 支払明細、LMS 実績テキストなど、形式の異なる売上関連資料から取得した情報を、共通の形式（以下「売上レコード / `sales_record`」）へ変換するためのデータモデルを定義する。

```
Excel 請求書 ──→ テンプレート別アダプター ─┐
PDF 支払明細 ──→ 支払明細アダプター ───────┼─→ sales_record ─→ 共通検証 ─→ 確認用 CSV ─→ 人間による確認
LMS テキスト ──→ LMS テキストアダプター ───┘
```

本書の範囲:

- `sales_record`（売上資料本体）の項目定義
- 日付・金額・null の扱い、検証ルール
- Human-in-the-loop のための状態管理
- MVP-1 で採用する項目と、将来拡張とする項目の区別

範囲外（概要のみ記載）:

- `bank_transaction`（銀行明細）、`payment_match`（入金照合結果）の詳細設計

## 2. 基本原則

1. **推測で値を埋めない。** 元資料から取得できない値、判断できない値は `null` とする。
2. **`0` と `null` を区別する。**
   - `0`: 元資料、または人間の確認により「該当なし・0 円」と確定している。
   - `null`: 元資料から取得できない、または判断できない。
   - 配列も同様に、`[]` は「該当なしと確定」、`null` は「不明」を表す。
3. **金額は円単位の整数（int）で扱う。** float は使用しない。税率等の計算が必要な場合は `Decimal` を用いる。
4. **ラベル名だけで項目へ機械的に割り当てない。** 資料上の「小計」「合計」「お支払額」等は、資料ごとに意味が異なる可能性がある。どのセル・どの行がどの項目に当たるかは、テンプレートごとに人間が確認して定義し、さらに金額の整合性検証で裏付ける。
5. **税務判断をプログラムが確定しない。** 売上計上日、源泉徴収の有無、消費税の扱い等、税務判断を伴う値は、資料に明記されたもの以外は人間が確定する。
6. **すべての出力は「候補」である。** プログラムが生成した `sales_record` は、人間が確認するまで確定扱いしない。
7. **実データの値をログ・警告メッセージに含めない。** 警告は「コード」と「項目名」で表現し、金額・取引先名等の実際の値を出力しない（CLAUDE.md ルール 19）。

## 3. `sales_record` 項目一覧

凡例:
- **型**: `str` / `int` / `date`（`YYYY-MM-DD` 形式の文字列）/ `enum` / `object` / `array`
- **null**: `可` = null を許容する
- **MVP-1**: ◎ = 採用 / ○ = 項目は持つが値は常に固定（例：常に null）/ – = 保留（将来拡張）

### 3.1 識別・メタ情報

| 項目 | 型 | null | MVP-1 | 説明 |
| --- | --- | --- | --- | --- |
| `schema_version` | str | 不可 | ◎ | データ形式のバージョン。初版は `"0.1"`。 |
| `record_id` | str | 不可 | ◎ | `sales_record` 自体を識別する内部 ID（UUID）。`source_hash` からは生成しない。§9 参照。 |
| `document_type` | enum | 不可 | ◎ | `invoice` / `payment_statement` / `lms_statement`。MVP-1 は `invoice` のみ。 |
| `source_type` | enum | 不可 | ◎ | `excel` / `pdf` / `lms_text` / `manual`（人間が手入力 CSV で記入した資料。年間処理で追加予定・未実装）。MVP-1 は `excel` のみ。 |
| `source_name` | str | 不可 | ◎ | 取引先・サービスを識別する内部名称（例：`template_a`）。§9.6 の注意を参照。 |
| `source_hash` | str | 不可 | ◎ | 元資料ファイル内容の SHA-256 ハッシュ値。同一ファイルの再取込・重複検出に使う。`record_id` とは独立。§9 参照。 |
| `external_id` | str | 可 | ○ | LMS 等の外部システム側の識別子。`record_id` / `source_hash` とは別概念。無い場合は null。MVP-1 では常に null。 |

### 3.2 資料の内容

| 項目 | 型 | null | MVP-1 | 説明 |
| --- | --- | --- | --- | --- |
| `document_date` | date | 可 | ◎ | 資料の日付。請求書の場合は **請求日**、支払明細の場合は発行日。§4 参照。 |
| `service_period` | object | 可 | ◎ | 業務対象期間。§4.2 参照。 |
| `customer` | str | 可 | ◎ | 取引先名（資料上の表記のまま）。 |
| `description` | str | 可 | ◎ | 業務内容の要約（資料上の件名等）。 |
| `line_items` | array | 可 | ◎ | 明細の配列。§6 参照。 |
| `payment_due` | date | 可 | ◎ | 支払期限。記載が無ければ null（月末等で補完しない）。 |

### 3.3 金額・税

| 項目 | 型 | null | MVP-1 | 説明 |
| --- | --- | --- | --- | --- |
| `net_amount` | int | 可 | ◎ | 税抜金額。 |
| `tax_amount` | int | 可 | ◎ | 消費税額。 |
| `gross_amount` | int | 可 | ◎ | 控除前の税込総額（源泉徴収・振込手数料等を差し引く前）。 |
| `tax_treatment` | enum | 不可 | ◎ | `exclusive`（外税）/ `inclusive`（内税）/ `none`（消費税なし）/ `unknown`（不明）。 |
| `tax_rate` | int | 可 | ◎ | 消費税率（パーセントの整数。例：`10`）。明細ごとに税率が異なる場合は null とし、明細側に持つ。 |
| `withholding_status` | enum | 不可 | ◎ | `none`（源泉徴収なし）/ `applied`（源泉徴収あり）/ `unknown`（不明）。§5.3 参照。 |
| `withholding_tax` | int | 可 | ◎ | 源泉徴収税額。§5.3 参照。 |
| `deductions` | array | 可 | ◎ | 源泉徴収 **以外** の控除。§5.4 参照。 |
| `expected_payment_amount` | int | 可 | ◎ | 実際に振り込まれると期待される金額。§5.5 参照。 |

### 3.4 税務判断・確認状態

| 項目 | 型 | null | MVP-1 | 説明 |
| --- | --- | --- | --- | --- |
| `revenue_date` | date | 可 | ○ | 売上計上日。**人間のみが設定する。** MVP-1 では常に null。§4.3 参照。 |
| `review_status` | enum | 不可 | ◎ | `unreviewed`（未確認）/ `needs_review`（要確認）/ `reviewed`（確認済み）。§8 参照。 |
| `warnings` | array | 不可 | ◎ | 検証で検出した警告の配列。無ければ `[]`。§8.2 参照。 |
| `calculated_fields` | array | 不可 | ◎ | 資料に記載が無く、プログラムが計算で導出した項目名の一覧。§8.3 参照。 |

### 3.5 将来拡張（MVP-1 では保留）

| 項目 | 説明 |
| --- | --- |
| `customer_id` | 取引先マスタへの参照。表記ゆれ（正式名称・略称・カナ振込名義）の吸収と入金照合で必要になる。 |
| `tax_breakdown` | 税率ごとの税抜金額・税額の内訳。税率が混在する資料に対応するため。 |
| `field_origins` | 項目ごとの値の由来（`source` / `calculated` / `human`）。§8.3 参照。 |
| `reviewed_by` / `reviewed_at` | 確認者・確認日時。 |
| `notes` | 人間による備考。 |

## 4. 日付の考え方

### 4.1 日付の概念の区別

以下は **別概念** として扱い、相互に補完・流用しない。

| 概念 | 項目 | 持つ場所 | 決め方 |
| --- | --- | --- | --- |
| 資料の日付（請求日・発行日） | `document_date` | `sales_record` | 資料に記載された値 |
| 業務対象期間 | `service_period` | `sales_record` | 資料に記載された値 |
| 支払期限 | `payment_due` | `sales_record` | 資料に記載された値 |
| 売上計上日 | `revenue_date` | `sales_record` | **人間が確定** |
| 入金日 | `payment_date` | `bank_transaction` / `payment_match`（将来） | 銀行明細の値 |

- `invoice_date` は独立項目として持たない。`document_type = invoice` の場合、`document_date` を請求日として扱う。
- `payment_date` は `sales_record` に持たせない。1 件の売上が複数回に分けて入金される場合や、複数の売上がまとめて入金される場合に、売上側に入金日を持つと表現できないため。

### 4.2 `service_period`

`from` / `to` / `month` を、それぞれ独立して null にできる構造とする。

```json
"service_period": {
  "from": "2026-02-01",
  "to": "2026-02-28",
  "month": null
}
```

| ケース | 表現 |
| --- | --- |
| 開始日・終了日とも判明 | `{"from": "2026-02-01", "to": "2026-02-28", "month": null}` |
| 開始日のみ判明 | `{"from": "2026-02-01", "to": null, "month": null}` |
| 「2026年2月分」のように月だけ判明 | `{"from": null, "to": null, "month": "2026-02"}` |
| 期間そのものが不明 | `"service_period": null` |

- `month` は `YYYY-MM` 形式の文字列。
- **月だけ判明している場合に、月初・月末の日付へ展開しない。**「2月分」が暦月を指すのか、締め日ベース（例：1/21〜2/20）なのかは資料からは判断できないため。
- `from` と `to` が両方ある場合は `from <= to` を検証する。

### 4.3 `revenue_date`

- 税務上の売上計上日は、請求日・業務期間・入金日のいずれとも一致するとは限らず、税務判断を伴う。
- **プログラムは `revenue_date` を設定しない。** `document_date` や `service_period` 等から自動的にコピー・推定しない。
- 人間が確認のうえで設定する。将来、候補を「提示」する機能を設ける場合も、候補は別項目として持ち、`revenue_date` そのものには入れない。

## 5. 金額

### 5.1 金額項目の定義

| 項目 | 意味 |
| --- | --- |
| `net_amount` | 税抜金額 |
| `tax_amount` | 消費税額 |
| `gross_amount` | 控除前の税込総額 |
| `withholding_tax` | 源泉徴収税額 |
| `deductions[].amount` | 源泉徴収以外の控除額（振込手数料等） |
| `expected_payment_amount` | 実際に振り込まれると期待される金額 |

- すべて円単位の整数。負の値は MVP-1 では扱わない（返金・値引き等は将来検討）。
- `gross_amount` は売上の総額であり、`expected_payment_amount` とは一致しないことがある。両者を混同しない。

### 5.2 `tax_treatment` と金額の関係

| `tax_treatment` | `net_amount` | `tax_amount` | `gross_amount` | 検証 |
| --- | --- | --- | --- | --- |
| `exclusive`（外税） | 資料の値 | 資料の値 | 資料の値 | `net_amount + tax_amount == gross_amount` |
| `inclusive`（内税） | 資料に記載があれば | 資料に記載があれば | 資料の値 | 3 項目とも判明している場合のみ同上 |
| `none`（消費税なし） | 資料の値 | `0` | 資料の値 | `net_amount == gross_amount` |
| `unknown`（不明） | 資料の値 | 資料の値 | 資料の値 | 判明している項目のみ検証し、`needs_review` とする |

- 内税で税額が資料に記載されていない場合、`tax_amount` は null とする（プログラムが逆算した値を資料の値として扱わない）。
- 内税で `net_amount` が資料に記載されていない場合に限り、以下が **すべて** 成立するときだけ、共通のレコード生成処理（アダプターではない）で `net_amount = gross_amount - tax_amount` を計算し、`calculated_fields` に `net_amount` を記録する（MVP-2 で実装）。
  - `tax_treatment` が `inclusive`
  - `net_amount` が null（資料の値を上書きしない）で、`net_amount` に読み取り時の警告が無い
  - `gross_amount` と `tax_amount` がともに null でない
  - `gross_amount >= tax_amount`
  - いずれかを満たさない場合は計算せず null のままとする。税額そのものは計算しない（資料に記載された 2 つの値の差を取るだけ）。
- `tax_treatment = none` は「資料上、消費税が明示的に 0 またはかからないと確定している」場合に限る。消費税の記載が無いだけの場合は `unknown` とする。

### 5.3 源泉徴収（`withholding_status` × `withholding_tax`）

| `withholding_status` | `withholding_tax` | 意味 |
| --- | --- | --- |
| `none` | `0` | 源泉徴収なしと確定（資料に明記、またはテンプレート定義で人間が確定済み） |
| `applied` | 正の整数 | 源泉徴収あり、金額判明 |
| `applied` | `null` | 源泉徴収ありだが金額不明 → `needs_review` |
| `unknown` | `null` | 源泉徴収の有無が不明 → `needs_review` |

上記以外の組み合わせ（例：`none` かつ `5000`、`unknown` かつ `0`）は不正として検証エラーとする。

- 資料に源泉徴収の記載が無いことだけをもって `none` としない。自分で作成した請求書に記載が無くても、支払側が源泉徴収する可能性があるため。
- テンプレート定義（§10）で、人間が「この取引先は源泉徴収なし」と確認済みの場合に限り、`none` を設定してよい。
- プログラムは源泉徴収税額を **計算しない**。資料に記載された値のみを使う。（対象となるかどうか・計算の基礎となる金額は税務判断を伴うため。）

### 5.4 `deductions`（源泉徴収以外の控除）

```json
"deductions": [
  {"type": "transfer_fee", "amount": 440, "description": "振込手数料（先方負担分を差引）"}
]
```

| 項目 | 型 | null | 説明 |
| --- | --- | --- | --- |
| `type` | enum | 不可 | `transfer_fee`（振込手数料）/ `offset`（相殺）/ `other`（その他） |
| `amount` | int | 不可 | 控除額（正の整数） |
| `description` | str | 可 | 補足説明 |

**源泉徴収との二重計上を防ぐルール:**

1. 源泉徴収税額は `withholding_tax` に **のみ** 記録する。
2. `deductions` に源泉徴収を含めてはならない（`type` に源泉徴収を表す値を用意しない）。
3. 資料上で源泉徴収と他の控除が合算表示されている場合は、内訳が判明しない限り分割せず、`needs_review` とする。

**`deductions` の範囲:**

- `sales_record.deductions` には、**資料に記載された、または事前に合意済みの控除** のみを記録する。
- `[]` は「資料上・合意上、控除なし」、`null` は「不明」を表す。
- 実際の入金時に初めて判明した差額（想定外の振込手数料等）は、売上レコードではなく、将来の `payment_match` 側で記録する。

### 5.5 `expected_payment_amount`

計算式:

```
expected_payment_amount = gross_amount - withholding_tax - sum(deductions[].amount)
```

- 資料に振込予定額（例：支払明細の「お支払額」）が記載されている場合は、その値を採用し、上記の式で整合性を検証する。不一致は警告とし、自動修正しない。
- 資料に記載が無い場合は、以下が **すべて** 確定している場合に限り計算し、`calculated_fields` に記録する。
  - `gross_amount` が null でない
  - `withholding_status` が `none` または `applied`（かつ `withholding_tax` が null でない）
  - `deductions` が null でない
- いずれかが不明な場合は null とする。

### 5.6 金額の検証ルール（共通）

| コード | 内容 | 重大度 |
| --- | --- | --- |
| `AMOUNT_NOT_INTEGER` | 円単位の整数として解釈できない金額値。非整数の数値（Excel の表示上は整数でも内部値が小数の場合を含む）や、金額として解釈できない文字列等を含む。値は丸めず null とする | error |
| `AMOUNT_NEGATIVE` | 金額が負 | error |
| `NET_TAX_GROSS_MISMATCH` | `net_amount + tax_amount != gross_amount`（§5.2 の条件下） | error |
| `LINE_ITEMS_SUM_MISMATCH` | 明細の合計が、`tax_treatment` が `exclusive`（外税）または `none`（消費税なし）の場合は `net_amount`、`inclusive`（内税）の場合は `gross_amount` と一致しない（`unknown` の場合は検証しない） | error |
| `LINE_ITEM_AMOUNT_MISMATCH` | 明細の `quantity × unit_price` が `amount` と一致しない（§6） | warning |
| `TAX_RATE_MISMATCH` | `tax_amount` が `net_amount × tax_rate` の切捨て・四捨五入・切上げのいずれとも一致しない | warning |
| `WITHHOLDING_INCONSISTENT` | §5.3 の組み合わせ表に無い組み合わせ | error |
| `EXPECTED_PAYMENT_MISMATCH` | 資料記載の振込予定額が §5.5 の計算式と一致しない | error |
| `AMOUNT_UNUSUALLY_LARGE` | 設定した上限額を超える | warning |

端数処理の方法そのものは税務・運用上の確認が必要なため、MVP-1 では「いずれかと一致すれば可」とし、特定の方法を正としない。

## 6. `line_items`

```json
"line_items": [
  {
    "description": "研修講師業務",
    "quantity": 2,
    "unit": "日",
    "unit_price": 15000,
    "amount": 30000,
    "tax_rate": null,
    "tax_treatment": null
  }
]
```

| 項目 | 型 | null | MVP-1 | 説明 |
| --- | --- | --- | --- | --- |
| `description` | str | 可 | ◎ | 明細の内容 |
| `quantity` | str | 可 | ◎ | 数量。小数（例：1.5 時間）があり得るため、float を避けて **文字列の十進数**（例：`"1.5"`）で保持し、計算時に `Decimal` へ変換する |
| `unit` | str | 可 | ◎ | 単位（例：日、時間、式） |
| `unit_price` | int | 可 | ◎ | 単価 |
| `amount` | int | 可 | ◎ | 明細金額 |
| `tax_rate` | int | 可 | – | 明細ごとの税率。null の場合はレコードの `tax_rate` に従う |
| `tax_treatment` | enum | 可 | – | 明細ごとの税の扱い。null の場合はレコードの値に従う |

- `line_items` が `null` の場合は「明細が取得できない」、`[]` の場合は「明細が無いと確定」を表す。
- `quantity × unit_price == amount` は、3 項目とも判明している場合に検証する（不一致は `LINE_ITEM_AMOUNT_MISMATCH`、warning。端数のある単価等があり得るため）。
- 立替経費（交通費等）が明細に含まれる場合の扱い（売上に含めるか等）は税務判断を伴うため、MVP-1 では区別せず、将来の検討事項とする。

## 7. 日付の検証ルール（共通）

| コード | 内容 | 重大度 |
| --- | --- | --- |
| `DATE_INVALID` | 実在しない日付（例：2月30日）、または解釈できない形式 | error |
| `DATE_AMBIGUOUS` | 年が無い等、一意に解釈できない（推測で補完しない） | warning |
| `DATE_OUT_OF_TARGET_YEAR` | `document_date` が対象年（設定値）の範囲外 | warning |
| `DUE_BEFORE_DOCUMENT` | `payment_due < document_date` | error |
| `DUE_TOO_FAR` | `payment_due` が `document_date` から設定日数を超えて離れている | warning |
| `PERIOD_REVERSED` | `service_period.from > service_period.to` | error |

和暦（例：令和8年）や `2026/2/28`、`2026年2月28日` 等の表記は、`YYYY-MM-DD` へ正規化する。

## 8. Human-in-the-loop

### 8.1 `review_status`

| 値 | 意味 | 設定者 |
| --- | --- | --- |
| `unreviewed` | 人間による確認前。検証上の問題が検出されていない状態。 | プログラム |
| `needs_review` | 警告、エラー、重要項目の欠落等が検出され、人間による確認が必要な状態。 | プログラム |
| `reviewed` | 人間による確認済み。 | **人間のみ** |

通常の流れ:

```
正常に抽出・検証 → unreviewed   → 人間が確認                   → reviewed
問題を検出       → needs_review → 人間が確認・必要に応じて修正 → reviewed
```

- `unreviewed` は「問題が無いと確定した」状態ではない。問題が検出されなかった場合でも、人間の確認を経て `reviewed` になるまでは確定扱いしない。
- プログラムは `reviewed` を設定しない。
- `warnings` に 1 件以上ある場合、または必須項目（`document_date`、`customer`、`gross_amount`）のいずれかが null の場合は `needs_review` とする。
- 確認済みでも警告は削除しない（「警告を認識したうえで確認した」ことが分かるようにする）。

### 8.2 `warnings`

```json
"warnings": [
  {"code": "TAX_RATE_MISMATCH", "field": "tax_amount", "severity": "warning"},
  {"code": "FIELD_MISSING", "field": "payment_due", "severity": "warning"}
]
```

| 項目 | 説明 |
| --- | --- |
| `code` | 警告コード（§5.6、§7、§11 等で定義） |
| `field` | 対象項目名（無ければ null） |
| `severity` | `error`（明らかな不整合）/ `warning`（要確認） |

- **警告に実データの値（金額・取引先名・日付等）を含めない。** 人間は確認用 CSV で値を見ながら、コードと項目名で内容を判断する。
- 項目が取得できなかった場合は `FIELD_MISSING` を記録する（欠落一覧の別項目は設けない）。
- `FIELD_MISSING`（warning）は、MVP-1 では次の場合に記録する。
  - 次の項目が null の場合：`document_date`、`service_period`、`customer`、`description`、`line_items`、`payment_due`、`net_amount`、`gross_amount`、`deductions`、`expected_payment_amount`。
  - `tax_amount` が null の場合（`tax_treatment` が `inclusive` のときを除く。§5.2）。
  - `tax_rate` が null の場合（`tax_treatment` が `exclusive` のときのみ）。
  - 明細の各項目（`description` / `quantity` / `unit` / `unit_price` / `amount`）が null の場合（`field` は `line_items[0].quantity` のように表す）。
  - `tax_treatment == "unknown"` の場合（`field` = `tax_treatment`）。
  - `withholding_status == "unknown"` の場合（`field` = `withholding_status`）。
  - `withholding_status == "applied"` かつ `withholding_tax` が null の場合（`field` = `withholding_tax`）。
  - ただし、同じ項目に既に別の警告（例：`FORMULA_NO_CACHED_VALUE`、`DATE_INVALID`）がある場合は重複して記録しない。
  - `external_id`・`revenue_date` は MVP-1 では常に null だが、`FIELD_MISSING` の対象としない。
- `unknown` を `FIELD_MISSING` として扱うのは、`unknown` を 0 や `none` とみなさず、人間による確認が必要な状態として `needs_review` に回すためである（§2 の null / 0 の区別、§5.3 の組み合わせ表は変わらない）。

### 8.3 値の由来（`source` / `calculated` / `human`）

**提案: MVP-1 では簡略版の `calculated_fields` のみを採用し、項目ごとの完全な由来管理（`field_origins`）は将来拡張とする。**

理由:

- MVP-1 では、プログラムは原則として資料の値をそのまま取得するだけで、補完・推定を行わない。計算で導出するのは `expected_payment_amount`（§5.5）、MVP-2 で追加した内税の `net_amount`（§5.2）等の限られた項目のみであり、それを `calculated_fields` に列挙すれば十分区別できる。
- 人間による修正は、MVP-1 では次の §8.4 の「ファイルを分ける」方式で区別できる。
- 項目ごとの由来管理は、人間の修正結果をシステムへ取り込む機能（MVP-1 の範囲外）と同時に設計するほうが、手戻りが少ない。

将来の `field_origins` の想定:

```json
"field_origins": {
  "gross_amount": "source",
  "expected_payment_amount": "calculated",
  "revenue_date": "human"
}
```

### 8.4 抽出結果と人間の確認結果の区別（MVP-1）

MVP-1 では、データ構造を複雑にせず、**ファイルを分けることで区別する**。

1. プログラムは抽出結果（`sales_record` の JSON）と確認用 CSV を出力する。
2. 抽出結果の JSON は **上書きしない**。
   - 同じ `source_hash` の資料が既存の出力に存在する場合は、新しい JSON を作成せず「既存取込候補」としてスキップする（§9.4）。
   - それ以外の資料は、`record_id` ごとに新しい JSON ファイルとして出力する。既存の JSON ファイルを上書きすることはない。
3. 人間は確認用 CSV をコピーし、確認・修正済みのファイルとして別名で保存する。
4. 確認済みファイルをシステムへ取り込む処理は、MVP-1 の範囲外（将来、`field_origins` と合わせて設計する）。

### 8.5 確認用 CSV（MVP-1）

- 出力先は `data/` 配下（Git 管理対象外）とする。実データを含むため。
- 1 行 = 1 `sales_record` のファイルと、1 行 = 1 明細のファイルの 2 つに分ける（`record_id` で対応付け）。
- `warnings` は各警告を `code:field` の形式（例：`FIELD_MISSING:payment_due`）で表し、複数ある場合は `;` で連結して 1 列に表示する。
- 人間が記入する列として `human_review_status`、`human_revenue_date`、`human_reviewer_note` を空欄で用意する。
  - プログラムが生成した値（`review_status` 等の列）と、人間が入力する確認値（`human_` で始まる列）は別の列とし、人間がプログラムの生成値を CSV 上で直接上書きする形にはしない。
  - `human_revenue_date` は人間が確定する売上計上日の入力欄であり、プログラムは値を入れない（`revenue_date` は §4.3 のとおりプログラムが設定しない）。
- 文字コードは Excel で開きやすい UTF-8（BOM 付き）を想定する（要確認）。
- **CSV インジェクション対策:** `=`、`+`、`-`、`@`、タブ、CR（復帰）で始まる文字列は Excel で数式として解釈されるおそれがあるため、出力時に先頭へ `'` を付けてエスケープする。数値型の値はエスケープしない。JSON 側の値はエスケープせず、資料の値のまま保持する。

## 9. 元資料の識別

### 9.1 方針

- ファイル名・ファイルパスを、共通データ・ログ・警告に保存しない。ファイル名に取引先名・個人名等が含まれる可能性があるため。
- 元資料は `source_hash`（ファイル内容の SHA-256）で識別する。

### 9.2 識別子の役割

`record_id`、`source_hash`、`external_id` は、それぞれ独立した別概念として扱う。

| 項目 | 何を識別するか | 生成・取得方法 | 主な用途 |
| --- | --- | --- | --- |
| `record_id` | `sales_record` 自体 | プログラムが UUID を生成 | 確認用 CSV・照合結果等からの参照 |
| `source_hash` | 元資料ファイル | ファイル内容の SHA-256 | 同一ファイルの再取込・重複の検出 |
| `external_id` | 外部システム上の対象 | 元資料（LMS 等）に記載された ID | 外部システムとの対応付け |

### 9.3 `record_id`

- `sales_record` 自体を識別する内部 ID とする。
- MVP-1 では UUID（例：UUID v4）で生成する。
- **`source_hash` や資料の内容から生成しない。** 元資料と売上レコードの識別を分けることで、元資料の修正・1 ファイルに複数の売上が含まれる場合等にも、レコードの識別が影響を受けない。

### 9.4 `source_hash` と再取込・重複の検出

- 元資料ファイル内容の SHA-256 とする。
- 同じ元ファイルを再処理した場合、`source_hash` が既存の取込済みレコードと一致することで **既存取込候補** として検出し、**自動的に別の売上として登録しない**（人間の確認に回す）。
  - MVP-1 では、CLI が出力先の既存 JSON の `source_hash` と照合し、一致した資料は `sales_record` を作成せずにスキップする。画面には既存の `record_id` を表示する。同一実行内で同じ資料が複数指定された場合もスキップする。
  - この検出は CLI の処理であり、`sales_record` の `warnings` に警告を追加するものではない。
- 1 ファイル（1 ブック）に複数の請求書が含まれる場合、それらのレコードは同じ `source_hash` を持つ。
- ファイルの内容が変わると `source_hash` も変わるため、修正版の請求書は `source_hash` では検出できない。
  - 将来の検討事項として、`document_date` + `customer` + `gross_amount` 等による内容ベースの重複候補検出（警告コード案：`DUPLICATE_CANDIDATE`）を想定している。
  - **MVP-1 では未実装** であり、`DUPLICATE_CANDIDATE` は実装済みの警告コードではない。

### 9.5 `external_id`

- LMS 等の外部システム側で付与された識別子とする。
- `record_id` / `source_hash` の代わりには使わない。外部 ID が無い資料もあり、また外部システム側の都合で変わる可能性があるため。
- 無い場合は null（MVP-1 では常に null）。

### 9.6 `source_name` とテンプレート名の注意

- `source_name` やテンプレート定義の名前は、公開リポジトリのコード・テストに現れる。
- **公開されるコード・テスト・sample_data では、実在の取引先名・サービス名を使わない**（例：`template_a`、`client_x` 等の汎用名を使う）。
- 実在の取引先とテンプレートの対応表は、`data/` 配下等の Git 管理対象外の場所に置く。

## 10. テンプレート別アダプター（MVP-1 / MVP-2）

「テンプレート別アダプター → 共通モデル」という構造で実装済みである。現在のアダプターは、いずれも完全な架空の書式に対するもの。

| テンプレート | 追加段階 | 書式の性質 | 税の扱い（既定値） | 主な割り当て |
| --- | --- | --- | --- | --- |
| `template_a` | MVP-1 | 外税型 | `exclusive`・10% | 税抜金額・消費税額・税込総額・源泉徴収税額・振込予定額を資料から取得 |
| `template_b` | MVP-2 | 内税・源泉徴収型 | `inclusive`・10% | 帳票上の「小計」を `gross_amount`、「（内消費税）」を `tax_amount` として取得。税抜金額の欄は無く、`net_amount` は §5.2 の条件で計算 |

- テンプレートは CLI の `--template` で明示的に指定する（既定は `template_a`）。テンプレートの自動判定は行わない。
- 同じラベル文字列（例：「小計」）でも書式ごとに意味が異なるため、ラベルから項目を推測せず、各テンプレート定義のセル位置で割り当てる。ラベルは書式の確認（`TEMPLATE_LABEL_MISMATCH`）にのみ使う。

```
完全な架空 Excel 請求書（Python スクリプトで生成）
  ↓
テンプレート定義（どのセル / ラベル位置がどの項目か。人間が定義）
  ↓
Excel 用アダプター（openpyxl で読み取り → sales_record へ変換）
  ↓
共通検証（§5.6、§7、§11）
  ↓
確認用 CSV 出力（§8.5）
  ↓
人間が確認
```

- アダプターは「資料の読み取りと項目への割り当て」だけを担当し、検証は共通処理で行う。これにより、MVP-2 以降でテンプレートや入力元（PDF・LMS テキスト）を追加しても、検証処理を共有できる。
- テンプレート定義には、セル位置・ラベル位置のほか、人間が確認済みの既定値（例：この取引先の請求書は外税・10%、源泉徴収なし）を持てるようにする。既定値はテンプレート定義者（人間）の判断として扱う。

## 11. Excel 読み取りの注意事項

### 11.1 数式セル

openpyxl は **数式を計算しない**。`data_only=True` で取得できる値は、Excel 等が最後に保存したときのキャッシュ値であり、以下の可能性がある。

- キャッシュ値が存在しない（Excel 以外のツールで作成・保存されたファイル等）→ 値が `None` になる
- キャッシュ値が古い（数式の参照先を変更後、再計算されずに保存された等）

このため、以下を方針とする。

- **数式セルの値を、無条件に正しい値として扱わない。**
- 数式セルのキャッシュ値が無い場合は null とし、`FORMULA_NO_CACHED_VALUE` 警告を記録する（推測で再計算した値を資料の値として扱わない）。
- **金額の整合性を Python 側でも必ず検証する**（§5.6）。キャッシュ値が古い場合も、合計の不一致として検出できる。

### 11.2 元ファイルの保護

- **元の Excel ファイルは読み取り専用として扱う。**
- **openpyxl で実会計 Excel を上書き保存しない。** openpyxl で保存すると、数式のキャッシュ値・書式・グラフ等が失われる可能性がある。
- 読み取り処理は、元ファイルを変更しない形で行う。

### 11.3 その他の注意

| 項目 | 注意点 | 対応 |
| --- | --- | --- |
| 表示値と内部値 | 表示上は `3,000` でも内部値が `3000.4` の場合がある | `AMOUNT_NOT_INTEGER` として警告 |
| 日付セル | 日付型のセルと、文字列で入力された日付が混在し得る | 両方を解釈し、解釈できなければ `DATE_INVALID` |
| 結合セル | 値は結合範囲の左上セルにのみ入っている | テンプレート定義で左上セルを指定 |
| 非表示シート・行 | 古い請求書や作業用の値が残っている場合がある | テンプレート定義で対象シートを明示 |
| ファイル形式 | 旧形式 `.xls` は openpyxl で読めない | 対象外とし、`.xlsx` への変換は人間が行う |

### 11.4 読み取り時の警告コード

| コード | 内容 | 重大度 |
| --- | --- | --- |
| `FORMULA_NO_CACHED_VALUE` | 数式セルにキャッシュ値が無い（§11.1） | warning |
| `TEMPLATE_LABEL_MISMATCH` | テンプレート定義で確認するラベルセルの内容が、定義と一致しない（別書式の誤適用の可能性） | error |
| `QUANTITY_INVALID` | 明細の数量を十進数として解釈できない | warning |

## 12. 入金照合との関係（概要のみ）

売上データと銀行明細は別データとして扱い、照合結果を独立させる。

```
sales_record ──┐
               ├──→ payment_match（照合結果・候補）
bank_transaction ┘
```

- `bank_transaction`: 銀行明細の 1 取引。入金日（`payment_date`）、金額、振込名義等を持つ。
- `payment_match`: `sales_record` と `bank_transaction` の対応。1 件ごとに充当金額を持つことで、以下に対応できる構造を想定する。
  - 1 売上 : 1 入金
  - 複数売上 : 1 入金（例：2 か月分がまとめて振り込まれる）
  - 1 売上 : 複数入金（例：分割入金）
- 入金時に判明した差額（想定外の振込手数料等）とその理由、照合の確認状態は `payment_match` 側で持つ。
- 詳細設計は MVP-5 / MVP-6 で行う。

## 13. データ例（完全な架空データ）

外税・源泉徴収あり・控除なしの請求書の例。

```json
{
  "schema_version": "0.1",
  "record_id": "00000000-0000-4000-8000-000000000001",
  "document_type": "invoice",
  "source_type": "excel",
  "source_name": "template_a",
  "source_hash": "0000000000000000000000000000000000000000000000000000000000000000",
  "external_id": null,
  "document_date": "2026-02-28",
  "service_period": {"from": null, "to": null, "month": "2026-02"},
  "customer": "株式会社サンプル",
  "description": "2026年02月研修業務委託費",
  "line_items": [
    {
      "description": "研修講師業務",
      "quantity": "2",
      "unit": "日",
      "unit_price": 15000,
      "amount": 30000,
      "tax_rate": null,
      "tax_treatment": null
    }
  ],
  "net_amount": 30000,
  "tax_amount": 3000,
  "gross_amount": 33000,
  "tax_treatment": "exclusive",
  "tax_rate": 10,
  "withholding_status": "applied",
  "withholding_tax": 3063,
  "deductions": [],
  "expected_payment_amount": 29937,
  "payment_due": "2026-03-31",
  "revenue_date": null,
  "review_status": "unreviewed",
  "warnings": [],
  "calculated_fields": ["expected_payment_amount"]
}
```

- 上記の金額・取引先名・`record_id`・ハッシュ値はすべて架空であり、源泉徴収税額は例示用の数値です。源泉徴収税額の計算方法を示すものではありません。
- `revenue_date` はプログラムが設定しないため null です。

## 14. MVP-1 で採用する項目・保留する項目（まとめ）

**採用（値を取得・検証する）:**
`schema_version`、`record_id`、`document_type`（invoice のみ）、`source_type`（excel のみ）、`source_name`、`source_hash`、`document_date`、`service_period`、`customer`、`description`、`line_items`（description / quantity / unit / unit_price / amount）、`net_amount`、`tax_amount`、`gross_amount`、`tax_treatment`、`tax_rate`、`withholding_status`、`withholding_tax`、`deductions`、`expected_payment_amount`、`payment_due`、`review_status`、`warnings`、`calculated_fields`

**項目は持つが MVP-1 では値が固定:**
`external_id`（常に null）、`revenue_date`（常に null。人間が設定）

**保留（将来拡張）:**
`customer_id`、`tax_breakdown`、明細ごとの `tax_rate` / `tax_treatment`、`field_origins`、`reviewed_by` / `reviewed_at`、`notes`、確認済み CSV の取り込み、`bank_transaction`、`payment_match`

## 15. 未決事項（人間の判断が必要）

1. 消費税の前提（課税事業者か免税事業者か、インボイス登録の有無、税率混在の有無、端数処理の方法）。
2. 源泉徴収の対象となる取引の範囲と、請求書上の記載方法。
3. 振込手数料の負担に関する取引先ごとの取り決め（`deductions` に事前に記録するか）。
4. 立替経費（交通費等）を明細に含める場合の扱い。
5. 返金・値引き・赤伝（負の金額）の扱い。
6. 確認用 CSV の文字コード（UTF-8 BOM 付きで Excel から問題なく開けるか）。
7. Excel と送付済み PDF の内容が食い違った場合に、どちらを正とするか。
8. 対象年・金額上限・支払期限の許容日数などの設定値。

## 16. 年間処理（全件再生成方式）のための追加定義

2026 年の実会計資料を処理する年間処理（`python -m blue_return_ai.annual`）のための定義。**現時点で実装済みなのは、実行単位の作成・資料一覧・`source_key`・売上 Excel（template_c 等）／カード CSV／銀行 CSV の取込と要確認一覧まで**（§17）であり、照合、仕訳、集計は未実装。

### 16.1 全件再生成と実行単位

- 年間処理は追記型ではなく、実行のたびに `data/` 配下の全資料から作り直す。
- 出力先は `data/output/annual/<run-id>/`。`run-id` は「対象年_実行日時_乱数」で一意にし、既存の出力先は上書きしない。
- 人間の判断は出力とは別に `data/decisions/` に保存し、再生成のたびに読み直して適用する。

### 16.2 `source_key`（実行をまたいで安定するキー）

| 項目 | 内容 |
| --- | --- |
| 目的 | 人間の判断ファイルと、資料・取引を実行をまたいで紐づける |
| `record_id` との違い | `record_id` は実行ごとに生成する UUID。`source_key` は同じ資料・同じ位置なら常に同じ値 |
| 形式 | `<namespace>-<キーの版>-<SHA-256 の先頭 32 文字>`（例：`sales-1-…`） |
| namespace | `sales` / `bank` / `card` / `household` / `document` |
| 売上資料 | `sales_source_key(source_hash, position)`：元ファイルの SHA-256 と資料内の位置（1 始まり）から作る |
| 個人情報 | 構成要素をハッシュ化するため、キーに個人情報そのものは含まれない |

- 関数は `src/blue_return_ai/keys.py` に実装済み。`sales_record` への項目追加は、年間処理で売上を取り込む段階で行う（既存の template_a / template_b の出力は変更していない）。
- 銀行・カード取引のキーは §17.4 を参照。

### 16.3 `documents_inventory`（資料一覧）

年間処理の実行ごとに、設定した各フォルダの全ファイルを一覧化する（ファイルの中身は解析しない）。

| 列 | 内容 |
| --- | --- |
| `document_key` | `document-` ＋ SHA-256 の先頭 16 文字 |
| `category` | `invoices` / `bank` / `credit_card` / `household` |
| `extension` | 拡張子 |
| `size_bytes` | ファイルサイズ |
| `sha256` | ファイル内容の SHA-256 |
| `status` | 下表 |
| `status_source` | `program`（プログラムが決めた）/ `human`（`data/decisions/documents.csv` で人間が指定） |
| `duplicate_of` | 同じ内容のファイルが先に出現していれば、その `document_key` |
| `relative_path` | リポジトリ直下からの相対パス（実ファイル名を含むため、出力は `data/` 配下のみ。共有用の表示・manifest には含めない） |
| `note` | 補足（例：`temporary_file`） |

| `status` | 意味 |
| --- | --- |
| `imported` | 取込処理で取り込んだ |
| `manual` | 手入力 CSV で対応した |
| `cross_check_only` | 照合用の資料（例：Excel と同じ請求の PDF）。売上等には使わない |
| `unsupported` | 取込対象外（対象外の拡張子、Excel の一時ファイル等） |
| `unresolved` | 取込対象だが未処理・要対応（PDF、テンプレートに合わない Excel、書式の合わない CSV、読み取りに失敗した資料、取込処理を設定していない種類の資料） |

- 取込処理が取り込むのは、プログラムが `unresolved` とした資料だけ。人間が `data/decisions/documents.csv` で状態を指定した資料は取り込まない。
- 同じ内容のファイル（`duplicate_of` があるもの）は取り込まず、`unsupported`（`note` = `duplicate_file`）にする。
- `note` には取り込めなかった理由（`template_sheet_not_found` / `format_mismatch` / `read_failed`）を記録する。

### 16.4 照合結果の状態（設計・未実装）

売上と銀行入金の照合など、プログラムが一致候補を出す処理では、次の状態を区別する。

| 状態 | 意味 | 集計での扱い |
| --- | --- | --- |
| `confirmed` | 人間が確認済み | 確定値に使う |
| `auto_candidate` | プログラム上は一致候補だが、人間は未確認 | 暫定の参考値として別欄に示す。確定値には混ぜない |
| `unresolved` | 不明・要確認 | 確定値から除外し、件数・金額を別欄に示す |

将来の `blue_return_summary` では、確定値・自動一致候補・要確認を区別して表示する。

## 17. 年間処理の取込（売上・カード・銀行）

設定ファイル（`data/config/annual.json`）の `processors` で、資料の種類ごとの取込方法を汎用の識別子で指定する。公開コードには金融機関名・サービス名・取引先名を書かない。

```json
"processors": {
  "invoices": {"templates": ["template_c", "template_instructor"]},
  "credit_card": {"format": "card_csv_a"},
  "bank": {"format": "bank_csv_a", "options": {"header_row": 13}}
}
```

- `invoices.template`（1 つ）または `invoices.templates`（複数）：`template_a` / `template_b` / `template_c` / `template_instructor`。複数の場合は、各テンプレートのシート名がブックにあるかだけで振り分ける（ファイル名・取引先名・金額は使わない）。該当するシートが無い Excel は `template_sheet_not_found`、複数のテンプレートのシートがある Excel は `template_ambiguous` として取り込まず `unresolved` のまま。シートが見つかった後のラベル構造の確認は各テンプレートの中で行う（合わなければ `TEMPLATE_LABEL_MISMATCH`）。PDF からは売上を作らない。
- `credit_card.format` / `bank.format`：CSV の書式。`options` で `encoding` / `header_row` / `columns` を上書きできる。
- 指定の無い種類は取り込まず、資料一覧だけを作る。

### 17.1 template_c（売上 Excel）

同じシート名で、明細見出しと集計欄のラベル構成が異なる 5 つのパターンが人間の構造調査で確認されている。パターンはファイル名・取引先名・金額では判定せず、ラベルの集合だけで判定する。

| パターン | 明細見出し（「品 番 • 品 名」以外） | 集計欄のラベル → 項目 | 税の扱い（テンプレート定義） |
| --- | --- | --- | --- |
| A | 数 量（日）／単 価／金 額 | 小計（税抜き）→ `net_amount`、消費税（10%）→ `tax_amount`、税込合計金額 → `gross_amount`、源泉徴収税額 → `withholding_tax`、振込依頼金額 → `expected_payment_amount` | `exclusive`・10% |
| B | 数 量（時間）／単 価／合計金額(内税) | 合計金額(税込み) → `gross_amount`、源泉徴収税(10.21%) → `withholding_tax`、振込金額(控除後支払額) → `expected_payment_amount`、報酬額(源泉対象) → 共通モデルに割り当てない（`withholding_base` として確認用に保持） | `inclusive`・税率は記載なし（null） |
| C | 数 量（時間）／単 価／金 額 | 小計 → `net_amount`、消費税（10%）→ `tax_amount`、合計金額 → `gross_amount` | `exclusive`・10% |
| D | 数 量／単 価／金額 | 小計 → 共通モデルに割り当てない（`document_subtotal` として確認用に保持） | `unknown` |
| E | 数 量（時間）／単 価／合計金額(内税)（B と同じ） | 合計(税込) → `gross_amount`、小計(税抜) → `net_amount`、消費税(内税10%) → `tax_amount`、源泉徴収税額 → `withholding_tax`、合計（差引支払額）→ `expected_payment_amount` | `inclusive`・10% |

判定の規則

- ラベルの比較は NFKC 正規化と空白除去の後の完全一致。全角・半角の括弧や空白の違いは同じとみなすが、文字が違えば別のラベル（例：「品番・品名」と「品番•品名」は別）。空白除去により「金 額」と「金額」は同じ文字列になるため、パターンは 1 つのラベルではなく見出し全体の組み合わせで区別する。
- 「品 番 • 品 名」がシート内に 1 つだけあり、その行のラベルの集合が、いずれかのパターンの明細見出しと完全に一致すること。
- 見出しより下にある「既知の集計ラベル（A〜E のいずれかに含まれるもの）」の集合が、そのパターンの集計ラベルと完全に一致し、各 1 つであること（不足・重複・他パターンのラベルの混在は不可）。
- 見出しと集計欄の両方が一致するパターンがちょうど 1 つの場合だけ採用する。B と E は明細見出しが同じだが、集計欄のラベル構成で区別し、別名として扱わない。
- 判定できない場合は推測せず `TEMPLATE_LABEL_MISMATCH`（`field` が null と `line_items`）とし、金額・明細を割り当てない。

値の取得

| 項目 | 取得方法 |
| --- | --- |
| `document_date` | 「請求日：」ラベル（全パターン必須）の右側。日付シリアル値は §17.1.1 |
| `customer` | `B3`（位置を確認済みのパターン A のみ。他は null） |
| 集計欄の各値・ご請求金額 | ラベルと同じ行で、ラベルより右にある最も近い空でないセル。値が無ければ null。型に合わなければ `AMOUNT_NOT_INTEGER` 等 |
| `line_items` | 見出しの次の行から、最も上の集計ラベルの直前まで。全項目が空の行は飛ばす。値は各見出しと同じ列。単位は見出しの「（日）」「（時間）」から（D は null） |
| 書式の確認 | パターン A のみ `B1` の表題も確認する |
| `service_period` / `description` / `payment_due` | この書式群には欄が無いため null（`FIELD_MISSING`） |

- 帳票に無い・意味が確認できない項目は null とし、0・`none`・計算値で埋めない。
  - B：`net_amount` / `tax_amount` / `tax_rate` は null（`FIELD_MISSING`）。内税の税抜金額の計算（§5.2）も、内消費税額が無いため行われない。
  - C：源泉徴収欄・支払額欄が無いため、`withholding_status` は `unknown`、`withholding_tax` / `expected_payment_amount` / `deductions` は null。
  - E：すべての金額が帳票の値（税抜金額も帳票の「小計(税抜)」であり計算しない）。整合は既存の検証（総額・明細合計・振込額）で確認する。
  - D：`net_amount` / `tax_amount` / `gross_amount` / `withholding_tax` / `expected_payment_amount` / `deductions` は null、`tax_treatment` は `unknown`。
- `deductions` は、振込額の欄があり控除欄が無いパターン（A・B・E）だけ `[]`（テンプレート定義）。
- 「ご請求金額」の値は `document_total` として確認用に保持する（共通モデルの金額には割り当てない）。
- 氏名欄など、売上データに不要な欄は読み取らない。
- 売上計上日（`revenue_date`）は設定しない。`sales.csv` では `revenue_date` を空欄、`revenue_date_status` を `unconfirmed` とし、要確認一覧に `REVENUE_DATE_UNCONFIRMED` を出す。
- `sales.csv` には、確認用 CSV の列（§8.5）に `source_key`・`document_key`・`revenue_date`・`revenue_date_status` と、確認用の値 `template_pattern`・`document_total`・`document_subtotal`・`withholding_base` を加える。`sales_record` 自体の項目は変更していない。

#### 17.1.2 template_instructor（シート「講師」）

template_c とは構造が大きく異なるため、独立したアダプターとする。

| 項目 | 取得方法 |
| --- | --- |
| 書式の確認 | シート「講師」。明細見出し「商品名」「数量」「単位」「単価(円)」「税率」「金額(円)」が同じ行に各 1 つ、確認済みの列（A / D / E / F / G / H）にあること。集計欄「小計」「消費税」「源泉税額」「合計」が見出しより下に各 1 つ、F 列にあること |
| `document_date` | 「請求日」ラベルの右側（日付セル・文字列・シリアル値。§17.1.1） |
| `line_items` | 見出しの次の行から、最初の集計欄・税率別内訳のラベルの直前まで。全項目が空の行は飛ばす。明細ごとの税率は「10%」・10・0.1 をパーセントの整数にする |
| `net_amount` / `tax_amount` / `withholding_tax` | 「小計」「消費税」「源泉税額」の右側の値 |
| 「合計」 | 源泉徴収控除後の支払額かを構造だけでは確定できないため、`expected_payment_amount` に割り当てず `document_total`（確認用）として保持 |
| `gross_amount` | 計算で作らない（null） |
| `tax_treatment` | 税率別内訳（「税率別内訳」「税抜金額」「消費税額」）がそろっていれば `exclusive`、なければ `unknown`。税率別内訳の数値は使わない |
| `withholding_status` | 源泉税額が 0 → `none`、正の値 → `applied`、空欄・読めない → `unknown` |
| `tax_rate`（記録全体） / `customer` / `description` / `payment_due` / `service_period` / `deductions` / `expected_payment_amount` | null（`FIELD_MISSING`） |

- 構造が確認できない場合は `TEMPLATE_LABEL_MISMATCH` とし、金額・明細を割り当てない。帳票パターン名は `instructor`（確認できない場合は `instructor_unidentified`）。

#### 17.1.1 Excel の日付シリアル値

- template_c の請求日セルが（表示形式が日付でない）数値の場合は、Excel の日付シリアル値として、openpyxl の `from_excel` でブックの基準日（1900 年基準 / 1904 年基準）に従って日付に変換する。
- 1 未満・9999-12-31 を超える値・真偽値・非有限値は `DATE_INVALID`。日付セル・文字列の日付は従来どおり。
- template_a / template_b の日付（`parse_date`）は変更していない（数値は従来どおり `DATE_INVALID`）。

### 17.2 カード取引（card_csv_a）

| 項目 | 内容 |
| --- | --- |
| `source_key` / `source_hash` / `document_key` / `row_number` | キー（§17.4）、元ファイルの SHA-256、資料一覧のキー、CSV 上の行番号 |
| `record_status` / `duplicate_of_document` | `active` / `duplicate`（§17.4）と、重複元の資料 |
| `use_date` | 利用日（経費計上日の候補）。支払月を経費の発生日として扱わない |
| `usage_category` / `user_category` / `merchant` / `point_target` / `installment_count` / `note` | 利用区分・利用者区分・利用店・ポイント対象・今回回数・備考（文字列のまま） |
| `amount` / `payment_amount` | 利用金額と今回の支払金額（別の項目。負の値も読み込む） |
| `business_classification` / `account_category` | `unknown` / null（利用店名から推測しない） |
| `review_status` / `warnings` | 警告があれば `needs_review` |

- ヘッダーより前の行は解析しない。ヘッダーが書式と一致しなければ資料全体を取り込まない（`CSV_HEADER_MISMATCH`）。
- 空行は飛ばす。列数の違う行・利用日を解釈できない行（合計行等）は取引にせず `ROW_NOT_TRANSACTION`。
- 警告：`CARD_NEGATIVE_AMOUNT`（負の利用金額）、`CARD_INSTALLMENT`（今回回数が 1 以外、または利用金額と支払金額が異なる）、`DATE_OUT_OF_TARGET_YEAR`、`FIELD_MISSING`、`AMOUNT_NOT_INTEGER`、`CLASSIFICATION_UNKNOWN`、`DUPLICATE_CANDIDATE`。

### 17.3 銀行取引（bank_csv_a）

| 項目 | 内容 |
| --- | --- |
| `source_key` / `source_hash` / `document_key` / `row_number` / `record_status` / `duplicate_of_document` | §17.2 と同じ |
| `sequence` | 明細通番（文字列のまま） |
| `transaction_date` | 日付（`YYYYMMDD` 形式にも対応） |
| `withdrawal_amount` / `deposit_amount` | 出金額・入金額（別の項目。空欄は null、0 は 0） |
| `balance` / `description` | 残高・取引内容 |
| `classification` / `account_category` | `unknown` / null（摘要から推測しない） |
| `review_status` / `warnings` | 警告があれば `needs_review` |

- 警告：`BANK_AMOUNT_MISSING`（出金・入金とも空欄）、`BANK_BOTH_AMOUNTS`（両方に 0 以外の値）、`AMOUNT_NEGATIVE`、`FIELD_MISSING`（残高）、`DATE_OUT_OF_TARGET_YEAR`、`AMOUNT_NOT_INTEGER`、`CLASSIFICATION_UNKNOWN`、`DUPLICATE_CANDIDATE`。

### 17.4 取引の `source_key` と重複除去

| 種類 | キーの材料（すべて SHA-256 でハッシュ化） |
| --- | --- |
| カード | 利用日・利用店・利用金額・支払金額・今回回数・利用区分・利用者区分 ＋ ファイル内で同じ内容が何回目か |
| 銀行 | 日付・出金額・入金額・残高・取引内容 ＋ ファイル内で同じ内容が何回目か |

- ファイルの SHA-256 はキーに含めない。期間の重なる別の CSV に同じ取引があれば同じキーになり、2 件目以降は `record_status` = `duplicate` として残す（集計・要確認の対象外）。
- 同じファイル内の、内容がまったく同じ正当な別取引は「何回目か」で区別し、消さない。
- 銀行の明細通番は、ダウンロードの単位ごとに振り直される可能性があるためキーに含めない。残高を含めるため、同日・同額・同摘要の別取引も通常は区別できる。
- 有効な取引のうち、主な内容（カード：利用日・利用店・利用金額、銀行：日付・出金額・入金額・取引内容）が同じでキーが異なるものには `DUPLICATE_CANDIDATE` を付け、消さずに人間の確認に回す。
- **限界**：公開コードに秘密値を持たないため、日付・金額・摘要の候補が分かる人は総当たりでキーとの一致を確かめられる。キーは `data/` 配下の出力にだけ保存し、外部に共有しない。また、取引の内容が後から変わった CSV（摘要の表記変更等）は別の取引として扱われる。

### 17.5 要確認一覧（`unresolved_items.csv`）

| 列 | 内容 |
| --- | --- |
| `item_id` | 実行内の連番 |
| `source_type` | `sales` / `card` / `bank` |
| `source_key` / `document_key` | 対象のキー（資料単位・行単位の問題は資料のキー） |
| `issue_code` / `severity` / `field` | 警告コード・重大度・対象項目（行単位の問題は `row <行番号>:<項目>`） |
| `message` | コードの説明（実データの値は含めない） |
| `decision_status` / `decision_value` / `decision_note` | 人間が判断を記入する列（空欄で出力） |

- 売上・有効な取引の警告、資料単位の問題（`FILE_READ_FAILED` / `CSV_HEADER_MISMATCH`）、取引として読まなかった行（`ROW_NOT_TRANSACTION`）、売上計上日の未確定（`REVENUE_DATE_UNCONFIRMED`）を 1 つにまとめる。
- 出力は実行ごとに作り直すため、判断を記入したファイルは別名で `data/decisions/` に保存する。判断ファイルを読み込んで反映する処理は次の段階で実装する。
