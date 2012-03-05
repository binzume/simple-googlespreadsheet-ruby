#!/usr/bin/ruby -Ku
require 'yaml'
require_relative 'spreadsheet'

conf = YAML.load_file('account.yaml')

spreadsheet_key = "0AlXPszXUpxhZdFdMRE1Rc1dEWjEtLXlqc3JDbms0Mmc"
account = {
  'accountType' => 'HOSTED_OR_GOOGLE',
  'Email' =>  conf['email'],
  'Passwd' => conf['passwd'],
  'service' => 'wise', # GAE:ah CAL:cl PLUS:oz SPREADSHEET:wise
  'source' => 'simple-spreadsheet.binzume.net',
}


# login
session = GSession.new
unless session.login(account)
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

