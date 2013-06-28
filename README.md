ruby-poi
========

Ruby bridge for Apache POI - the Java API for Microsoft Documents.

Making Office Document manipulations just a little bit less painful.

INSTALL
========
gem install ruby-poi

USAGE
=====

```ruby
require 'ruby-poi'
# Open Excel file.
workbook = POI::Workbook.open('spreadsheet.xlsx')
    
# Get to sheets, rows and columns
sheet    = workbook[0]
sheet    = workbook["Sheet 1"]
sheet    = workbook.worksheets[0]
sheet    = workbook.worksheets["Sheet 1"]
rows     = sheet.rows
columns  = sheet.columns

# Get values form sheets 
sheet[0][0]       # => Cell
sheet[0][0].value # => 1.0
sheet['A1']       # => Cell
sheet['A1'].value # => 1.0

# Get values form rows
rows[0][0]       # => Cell
rows[0][0].value # => 1.0
rows['A1']       # => Cell
rows['A1'].value # => 1.0
    
# Loop over cells.
workbook.each do |sheet|
  sheet.each do |row|
    row.each do |cell|
      puts cell.value
      if cell.value == 'FILL IT INN'
        cell.value = 'New Value'
      end
    end
  end
end
    
# Save Excel file.
workbook.save! # => Save it to 'spreadsheet2.xlsx'
workbook.save_as('spreadsheet2.xlsx')
```

Note on Patches/Pull Requests
=============================

* Fork the project.
* Make your feature addition or bug fix.
* Add tests for it. This is important so I don't break it in a future version unintentionally.
* Commit, do not mess with rakefile, version, or history.
  (if you want to have your own version, that is fine but bump version in a commit by itself I can ignore when I pull)
* Send me a pull request.

Copyright
=========

Copyright (c) 2010 Ægir Örn Símonarson.
See NOTICE and LICENSE for details.
