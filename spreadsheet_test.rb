#!/usr/bin/ruby -Ku
require_relative 'spreadsheet'

spreadsheet_key = "0AlXPszXUpxhZdFBLZXptOGlsOU9aRFpwaTVwTmtSekE"

# login (Get OAuth2 params from Google API Console (support only Service Account))
session = GSession.new
unless session.login_with_oauth2(JSON.parse(File.read('/root/binzume-bot-2f903d30a6df.json')))
  puts "auth error!"
  exit
end

# get spreadsheet
spread = GSpreadsheet.new(session, spreadsheet_key)
puts spread.title

# get worksheet
ws = spread.worksheets[0]

# puts meta data
puts ws.title
puts ws.row_count
puts ws.col_count

# get A2 value ( ws[row, col] )
p ws[2, 1]

# update A2
ws[2, 1] = "aihghaop"

data = [
  ["asdfghjk", "hoge", "12345"],
  ["0", "ma<>'aaa", "ag@aga"],
]

# fill A3,B3,C3 , A4,B4,C4
ws.set_cells(3,1,data)

