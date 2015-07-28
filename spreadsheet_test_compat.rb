#!/usr/bin/ruby -Ku
require 'json'
require_relative 'spreadsheet'

spreadsheet_key = "0AlXPszXUpxhZdFBLZXptOGlsOU9aRFpwaTVwTmtSekE"

# login
session = GoogleSpreadsheet.login_with_oauth2(JSON.parse(File.read('/root/binzume-bot-2f903d30a6df.json')))

# get first workseet
ws = session.spreadsheet_by_key(spreadsheet_key).worksheets[0]

# puts meta data
puts ws.title

# get A2 value ( ws[row, col] )
a2 = ws[2, 1]
p a2

# update A2
ws[2, 1] = a2.to_i + 1

# fill A3,B3,C3 , A4,B4,C4
data = [
  ["foo", "bar", "12345"],
  ["test", "ma<>'aaa", "ag@aga"],
]
ws.set_cells(3,1,data)

p ws[2]

# add row last
p ws.row_count
ws << ["hoge",1,2,3,4]

