# 共通売上データモデル 設計書

- 状態: **設計段階（未実装）**
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
| `source_type` | enum | 不可 | ◎ | `excel` / `pdf` / `lms_text`。MVP-1 は `excel` のみ。 |
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
| `AMOUNT_NOT_INTEGER` | 金額が整数でない（Excel の表示上は整数でも内部値が小数の場合を含む） | error |
| `AMOUNT_NEGATIVE` | 金額が負 | error |
| `NET_TAX_GROSS_MISMATCH` | `net_amount + tax_amount != gross_amount`（§5.2 の条件下） | error |
| `LINE_ITEMS_SUM_MISMATCH` | 明細の合計が `net_amount`（外税）または `gross_amount`（内税）と一致しない | error |
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
- `quantity × unit_price == amount` は、3 項目とも判明している場合に検証する（不一致は warning。端数のある単価等があり得るため）。
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

### 8.3 値の由来（`source` / `calculated` / `human`）

**提案: MVP-1 では簡略版の `calculated_fields` のみを採用し、項目ごとの完全な由来管理（`field_origins`）は将来拡張とする。**

理由:

- MVP-1 では、プログラムは原則として資料の値をそのまま取得するだけで、補完・推定を行わない。計算で導出するのは `expected_payment_amount` 等の限られた項目のみであり、それを `calculated_fields` に列挙すれば十分区別できる。
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
2. 抽出結果の JSON は **上書きしない**（再実行時は新しいファイルとして出力する）。
3. 人間は確認用 CSV をコピーし、確認・修正済みのファイルとして別名で保存する。
4. 確認済みファイルをシステムへ取り込む処理は、MVP-1 の範囲外（将来、`field_origins` と合わせて設計する）。

### 8.5 確認用 CSV（MVP-1）

- 出力先は `data/` 配下（Git 管理対象外）とする。実データを含むため。
- 1 行 = 1 `sales_record` のファイルと、1 行 = 1 明細のファイルの 2 つに分ける（`record_id` で対応付け）。
- `warnings` はコードを連結して 1 列に表示する。
- 人間が記入する列（例：`review_status`、`revenue_date`、`reviewer_note`）を空欄で用意する。
- 文字コードは Excel で開きやすい UTF-8（BOM 付き）を想定する（要確認）。
- **CSV インジェクション対策:** `=`、`+`、`-`、`@` で始まる文字列は Excel で数式として解釈されるおそれがあるため、文字列項目は出力時にエスケープする。

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
- 1 ファイル（1 ブック）に複数の請求書が含まれる場合、それらのレコードは同じ `source_hash` を持つ。
- ファイルの内容が変わると `source_hash` も変わるため、修正版の請求書は `source_hash` では検出できない。このため、`document_date` + `customer` + `gross_amount` 等による **重複候補の検出** を併用する（`DUPLICATE_CANDIDATE` 警告）。

### 9.5 `external_id`

- LMS 等の外部システム側で付与された識別子とする。
- `record_id` / `source_hash` の代わりには使わない。外部 ID が無い資料もあり、また外部システム側の都合で変わる可能性があるため。
- 無い場合は null（MVP-1 では常に null）。

### 9.6 `source_name` とテンプレート名の注意

- `source_name` やテンプレート定義の名前は、公開リポジトリのコード・テストに現れる。
- **公開されるコード・テスト・sample_data では、実在の取引先名・サービス名を使わない**（例：`template_a`、`client_x` 等の汎用名を使う）。
- 実在の取引先とテンプレートの対応表は、`data/` 配下等の Git 管理対象外の場所に置く。

## 10. MVP-1: テンプレート別アダプター

MVP-1 は「テンプレート別アダプター → 共通モデル」という構造を前提とする（コードは今回作成しない）。

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
