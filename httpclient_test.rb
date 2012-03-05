#!/usr/bin/ruby -Ku

require_relative 'httpclient'

client = HTTPClient.new

r = client.get("http://www.binzume.net/")
puts r.code
puts r['Content-Type']
puts r.body
