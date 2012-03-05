#!/usr/bin/ruby -Ku
require 'yaml'
require_relative 'spreadsheet'

conf = YAML.load_file('account.yaml')

spreadsheet_key = "0AlXPszXUpxhZdFdMRE1Rc1dEWjEtLXlqc3JDbms0Mmc"

# login
session = GoogleSpreadsheet.login(conf['email'], conf['passwd'])

# get first workseet
ws = session.spreadsheet_by_key(spreadsheet_key).worksheets[0]

# puts meta data
puts ws.title

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

