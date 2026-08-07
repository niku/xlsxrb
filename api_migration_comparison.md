# xlsxrb マイグレーション & API比較ガイド

既存の主要なRuby向けExcelライブラリ（`caxlsx`, `xlsxtream`, `roo` など）の典型的な使われ方を、`xlsxrb` で書き直すとどうなるかを比較しました。

## 1. 書き込み（In-Memory）: `caxlsx` (axlsx) vs `xlsxrb`

`caxlsx` は最も人気のある書き込みライブラリですが、DSLが少し古く、シリアライズ処理を手動で呼ぶ必要があります。

### Before: `caxlsx`
GitHubでよく見かける標準的なエクスポート処理です。
```ruby
require 'axlsx'

Axlsx::Package.new do |p|
  p.workbook.add_worksheet(name: "Sales Data") do |sheet|
    # ヘッダーの追加
    sheet.add_row ["Date", "Amount", "Status"]
    
    # スタイルの適用（事前定義が必要）
    style = p.workbook.styles.add_style(bg_color: "FF0000", fg_color: "FFFFFF")
    
    # データの追加
    sheet.add_row [Date.today, 1000, "Pending"], style: [nil, nil, style]
  end
  p.serialize('sales.xlsx')
end
```

### After: `xlsxrb`
`xlsxrb` はより直感的で、ブロックを抜けると自動でファイルが保存されます。スタイルもインラインで柔軟に書けます。
```ruby
require 'xlsxrb'

Xlsxrb.generate("sales.xlsx") do |wb|
  wb.sheet("Sales Data") do |sheet|
    # ヘッダーの追加
    sheet.row(["Date", "Amount", "Status"])
    
    # スタイルの適用（インラインで直感的に記述可能）
    sheet.row([Date.today, 1000, "Pending"]) do |r|
      r.cell(2).style(fill_color: "FF0000", font_color: "FFFFFF")
    end
  end
end
# ブロック終了時に自動でファイル生成・クローズされる
```

---

## 2. 書き込み（Streaming）: `xlsxtream` vs `xlsxrb`

大量データを省メモリで書き出す `xlsxtream` の比較です。

### Before: `xlsxtream`
```ruby
require 'xlsxtream'

Xlsxtream::Workbook.open('huge_data.xlsx') do |xlsx|
  xlsx.write_worksheet 'Users' do |sheet|
    sheet << ['ID', 'Name', 'Email']
    100_000.times do |i|
      sheet << [i, "User #{i}", "user#{i}@example.com"]
    end
  end
end
```

### After: `xlsxrb`
`xlsxrb` の `generate` もデフォルトでストリーミング動作するため、ほぼ同じ書き心地で移行でき、かつスタイルなどの高度な機能も同時に使えます。
```ruby
require 'xlsxrb'

Xlsxrb.generate("huge_data.xlsx") do |wb|
  wb.sheet("Users") do |sheet|
    sheet.row(['ID', 'Name', 'Email'])
    100_000.times do |i|
      sheet.row([i, "User #{i}", "user#{i}@example.com"])
    end
  end
end
```

---

## 3. 読み込み（Streaming）: `roo` vs `xlsxrb`

データのインポートやバッチ処理で最もよく使われる `roo` のストリーミング読み込みの比較です。

### Before: `roo`
```ruby
require 'roo'

xlsx = Roo::Excelx.new("data.xlsx")
xlsx.sheet(0).each_row_streaming(pad_cells: true) do |row|
  # row は Roo::Excelx::Cell の配列
  values = row.map { |cell| cell ? cell.value : nil }
  puts values.join(", ")
end
```

### After: `xlsxrb`
`xlsxrb.foreach` を使えば、Rooよりもはるかに高速に、かつ低いメモリ消費で同じ処理が書けます。
```ruby
require 'xlsxrb'

Xlsxrb.foreach("data.xlsx") do |sheet|
  sheet.each do |row|
    # row.cells は Xlsxrb::Elements::Cell の配列
    values = row.cells.map(&:value)
    puts values.join(", ")
  end
end
```

---

## 使い勝手の総評

* **学習コスト:** 他のライブラリ（特に `caxlsx` や `xlsxtream`）のDSLに非常に近い直感的なインターフェース（`sheet` や `row` をブロックでネストする）を採用しているため、Rubyエンジニアであれば数分で使いこなせます。
* **書き心地の良さ:** `caxlsx` のように「事前定義したスタイルオブジェクトを配列で渡す」といった煩雑さがなく、ブロック内でインラインにスタイルや幅を指定できるため、Rubyらしく美しいコードになります。
* **統一感:** 従来は「書き出しは `caxlsx`、読み込みは `roo`」のように2つの異なるライブラリとそのDSLを覚える必要がありましたが、`xlsxrb` ひとつで両方を統一された記法で扱えるのが最大のメリットです。
