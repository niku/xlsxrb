# 実践的 API マイグレーション比較（複雑なユースケース 20選）

より現実に即した、業務アプリケーションでよく使われる高度な20のユースケースについて、他のライブラリのコードを `xlsxrb` に書き換える方法をまとめました。

---

## 🎨 1. 書き込み (caxlsx -> xlsxrb)

### 1. セルの結合とセンタリング
**Before (caxlsx):**
```ruby
sheet.add_row ["Q1 Results", "", ""]
sheet.merge_cells("A1:C1")
style = p.workbook.styles.add_style(alignment: { horizontal: :center })
sheet.rows.last.style = style
```
**After (xlsxrb):**
```ruby
sheet.row(["Q1 Results", nil, nil]) do |r|
  r.merge_cells(0..2)
  r.cell(0).style(align_horizontal: "center")
end
```

### 2. 条件付き書式 (Conditional Formatting)
**Before (caxlsx):**
```ruby
sheet.add_row [50, 150, 80]
sheet.add_conditional_formatting("A1:C1", { type: :cellIs, operator: :greaterThan, formula: "100", dxfId: red_style_id, priority: 1 })
```
**After (xlsxrb):**
```ruby
sheet.row([50, 150, 80])
sheet.conditional_format("A1:C1", type: "cellIs", operator: "greaterThan", formula: "100", dxf_id: red_style_id)
```

### 3. データの入力規則（ドロップダウンリスト）
**Before (caxlsx):**
```ruby
sheet.add_data_validation("B2:B10", { type: :list, formula1: '"Pending,Approved,Rejected"', showDropDown: false })
```
**After (xlsxrb):**
```ruby
sheet.validate_data("B2:B10", type: "list", formula1: '"Pending,Approved,Rejected"', show_drop_down: false)
```

### 4. 列幅の自動調整・手動指定
**Before (caxlsx):**
```ruby
sheet.add_row ["Very Long Text", "Short"]
sheet.column_widths 30, nil
```
**After (xlsxrb):**
```ruby
sheet.row(["Very Long Text", "Short"])
sheet.column(0, width: 30)
# または auto_fit も可能
```

### 5. シートの保護 (パスワード)
**Before (caxlsx):**
```ruby
sheet.sheet_protection.password = 'secret'
```
**After (xlsxrb):**
```ruby
sheet.protect_sheet(password: 'secret')
```

### 6. オートフィルタとウィンドウ枠の固定
**Before (caxlsx):**
```ruby
sheet.add_row ["ID", "Name"]
sheet.auto_filter = "A1:B10"
sheet.sheet_view.pane do |pane|
  pane.state = :frozen
  pane.y_split = 1
end
```
**After (xlsxrb):**
```ruby
sheet.row(["ID", "Name"])
sheet.auto_filter("A1:B1")
sheet.freeze_pane(row: 1)
```

### 7. ハイパーリンクの追加
**Before (caxlsx):**
```ruby
sheet.add_row ["Google"]
sheet.add_hyperlink location: "https://google.com", ref: sheet.rows.last.cells.first
```
**After (xlsxrb):**
```ruby
sheet.row(["Google"]) do |r|
  r.cell(0).hyperlink("https://google.com")
end
```

### 8. 通貨や日付のフォーマット
**Before (caxlsx):**
```ruby
currency = p.workbook.styles.add_style(format_code: '"$"#,##0.00')
sheet.add_row [1000], style: currency
```
**After (xlsxrb):**
```ruby
sheet.row([1000]) do |r|
  r.cell(0).style(num_fmt: '"$"#,##0.00')
end
```

### 9. グラフの追加 (棒グラフ)
**Before (caxlsx):**
```ruby
sheet.add_chart(Axlsx::Bar3DChart, start_at: "A5", end_at: "F15") do |chart|
  chart.add_series data: sheet["B2:B4"], labels: sheet["A2:A4"], title: sheet["B1"]
end
```
**After (xlsxrb):**
```ruby
sheet.add_chart(type: :bar, from: "A5", to: "F15") do |chart|
  chart.series(data: "B2:B4", categories: "A2:A4", title: "B1")
end
```

### 10. 印刷設定 (ヘッダー/フッター・余白)
**Before (caxlsx):**
```ruby
sheet.page_setup.set(orientation: :landscape)
sheet.header_footer.center_header = "Confidential"
```
**After (xlsxrb):**
```ruby
sheet.page_setup(orientation: "landscape")
sheet.header_footer(center_header: "Confidential")
```

---

## ⚡ 11-13. ストリーミング書き出し (xlsxtream -> xlsxrb)

### 11. ActiveRecord の大量レコードをバッチエクスポート
**Before (xlsxtream):**
```ruby
Xlsxtream::Workbook.open(file_path) do |xlsx|
  xlsx.write_worksheet 'Users' do |sheet|
    sheet << ['ID', 'Email', 'Created At']
    User.find_each(batch_size: 1000) do |user|
      sheet << [user.id, user.email, user.created_at]
    end
  end
end
```
**After (xlsxrb):**
```ruby
Xlsxrb.generate(file_path) do |wb|
  wb.sheet('Users') do |sheet|
    sheet.row(['ID', 'Email', 'Created At'])
    User.find_each(batch_size: 1000) do |user|
      sheet.row([user.id, user.email, user.created_at])
    end
  end
end
```

### 12. Rails コントローラーからの StringIO ダウンロード
**Before (xlsxtream):**
```ruby
io = StringIO.new
Xlsxtream::Workbook.new(io).write_worksheet 'Data' do |s|
  s << ['Data']
end
send_data io.string, filename: "data.xlsx"
```
**After (xlsxrb):**
```ruby
io = StringIO.new
Xlsxrb.generate(io) do |wb|
  wb.sheet('Data') { |s| s.row(['Data']) }
end
send_data io.string, filename: "data.xlsx"
```

### 13. 複数シートへの分割ストリーミング
**Before (xlsxtream):**
```ruby
xlsx.write_worksheet 'A' do |s| s << [1] end
xlsx.write_worksheet 'B' do |s| s << [2] end
```
**After (xlsxrb):**
```ruby
wb.sheet('A') { |s| s.row([1]) }
wb.sheet('B') { |s| s.row([2]) }
```

---

## 📖 14-17. 読み込み (roo -> xlsxrb)

### 14. ヘッダー行をキーにしたHash配列への変換
**Before (roo):**
```ruby
xlsx = Roo::Excelx.new("data.xlsx")
headers = xlsx.row(1)
data = (2..xlsx.last_row).map do |i|
  Hash[headers.zip(xlsx.row(i))]
end
```
**After (xlsxrb):**
```ruby
data = []
Xlsxrb.foreach("data.xlsx") do |sheet|
  headers = sheet.first.cells.map(&:value)
  sheet.each_with_index do |row, idx|
    next if idx == 0
    data << headers.zip(row.cells.map(&:value)).to_h
  end
end
```

### 15. 条件に一致する行だけの抽出 (Filter)
**Before (roo):**
```ruby
target_rows = []
xlsx.each_row_streaming do |row|
  target_rows << row if row[0]&.value == 'Target'
end
```
**After (xlsxrb):**
```ruby
target_rows = []
Xlsxrb.foreach("data.xlsx") do |sheet|
  sheet.each do |row|
    target_rows << row if row.cells[0]&.value == 'Target'
  end
end
```

### 16. セルからのハイパーリンク抽出
**Before (roo):**
```ruby
link = xlsx.excelx_value(1, 1).link
```
**After (xlsxrb):**
※ `xlsxrb` ではパース時にリンク情報もCellオブジェクトに付随します。
```ruby
Xlsxrb.foreach("data.xlsx") do |sheet|
  link = sheet.first.cells[0].hyperlink
end
```

### 17. 複数シートの読み込み
**Before (roo):**
```ruby
xlsx.sheets.each do |name|
  xlsx.sheet(name).each_row_streaming { |r| puts r }
end
```
**After (xlsxrb):**
```ruby
Xlsxrb.foreach("data.xlsx") do |sheet|
  puts "Sheet: #{sheet.name}"
  sheet.each { |r| puts r.cells.map(&:value) }
end
```

---

## 📝 18-20. 既存ファイルの編集 (rubyXL -> xlsxrb)

既存のファイルを開き、一部を書き換えて保存するパターンです。

### 18. 特定のセルのValueを上書きする
**Before (rubyXL):**
```ruby
workbook = RubyXL::Parser.parse("template.xlsx")
sheet = workbook[0]
sheet.add_cell(0, 0, 'Updated Value')
workbook.write("output.xlsx")
```
**After (xlsxrb):**
```ruby
Xlsxrb.modify("template.xlsx", "output.xlsx") do |wb|
  wb.sheet(0) do |sheet|
    sheet.cell("A1").value = 'Updated Value'
  end
end
```

### 19. テンプレート変数の置換 (Search & Replace)
**Before (rubyXL):**
```ruby
workbook = RubyXL::Parser.parse("invoice.xlsx")
workbook[0].each do |row|
  row&.cells&.each do |cell|
    if cell&.value == '{{CUSTOMER_NAME}}'
      cell.change_contents('John Doe')
    end
  end
end
**After (xlsxrb):**
```ruby
Xlsxrb.modify("invoice.xlsx", "invoice_out.xlsx") do |wb|
  wb.sheet(0) do |sheet|
    sheet.each_cell do |cell|
      if cell.value == '{{CUSTOMER_NAME}}'
        cell.value = 'John Doe'
      end
    end
  end
end
```

### 20. 既存ファイルへの新規シート追加
**Before (rubyXL):**
```ruby
workbook = RubyXL::Parser.parse("data.xlsx")
workbook.add_worksheet('New Sheet')
workbook.write("data_out.xlsx")
```
**After (xlsxrb):**
```ruby
Xlsxrb.modify("data.xlsx", "data_out.xlsx") do |wb|
  wb.add_sheet('New Sheet') do |sheet|
    sheet.row(["New", "Data"])
  end
end
```
